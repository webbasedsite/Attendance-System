function doPost(e) {
  try {
    const params = e.parameter;
    const action = getParam(params, 'action', true);

    const Roles = Object.freeze({
      AGENT: "agent",
      INCHARGE: "incharge",
      ADMIN: "admin"
    });

    const SHEET_ID = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
    if (!SHEET_ID) return jsonResponse(false, "Sheet ID not configured in Script Properties");

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const employeesSheet = ss.getSheetByName("Employees");
    const officesSheet = ss.getSheetByName("Offices");
    const attendanceSheet = ss.getSheetByName("Attendance");

    const employeesData = getSheetData(employeesSheet);
    const officesData = getSheetData(officesSheet);
    const attendanceData = getSheetData(attendanceSheet);

    const colIndexes = getColumnIndexes(employeesData.headers, ["Phone", "Password", "OfficeID", "Role", "Name"]);
    if (colIndexes.includes(-1)) return jsonResponse(false, "Employees sheet missing required columns");

    const [phoneCol, passwordCol, officeIdCol, roleCol, nameCol] = colIndexes;
    const rawPhone = getParam(params, 'phone', false);
    const normalizedPhone = rawPhone ? normalizePhone(rawPhone) : null;

    if (normalizedPhone && !["addEmployee", "getOffices", "getAllEmployees", "login"].includes(action)) {
      if (rateLimitExceeded(normalizedPhone)) {
        return jsonResponse(false, "Rate limit exceeded. Please wait.");
      }
    }

    Logger.log(`Received action: ${action} from phone: ${normalizedPhone || "N/A"}`);

    // -------- ADD EMPLOYEE --------
    if (action === "addEmployee") {
      const office = getParam(params, 'office');
      const name = getParam(params, 'name');
      const phoneNew = normalizePhone(getParam(params, 'phone'));
      const role = getParam(params, 'role');
      const password = getParam(params, 'password');
      const latitude = parseFloat(params.latitude);
      const longitude = parseFloat(params.longitude);
      const accuracy = parseFloat(params.accuracy);

      if (password.length < 6) return jsonResponse(false, "Password must be at least 6 characters");
      if (isNaN(latitude) || isNaN(longitude) || isNaN(accuracy)) return jsonResponse(false, "Location data required");
      if (accuracy > 50) return jsonResponse(false, "Location accuracy is too low (must be ≤ 50 meters)");

      const officeExists = officesData.rows.some(o => o[0] === office);
      if (!officeExists) return jsonResponse(false, "Office does not exist");

      const phoneExists = employeesData.rows.some(r => normalizePhone(r[phoneCol]) === phoneNew);
      if (phoneExists) return jsonResponse(false, "Employee with this phone number already exists");

      employeesSheet.appendRow([
        phoneNew,
        password,
        office,
        role,
        name
      ]);

      Logger.log(`Added new employee: ${name}, phone: ${phoneNew}`);
      return jsonResponse(true, "Employee added successfully");
    }

    // -------- LOGIN --------
    if (action === "login") {
      const phone = normalizePhone(getParam(params, 'phone'));
      const password = getParam(params, 'password');

      const matched = employeesData.rows.find(r =>
        normalizePhone(r[phoneCol]) === phone &&
        String(r[passwordCol]).trim() === password
      );

      if (!matched) return jsonResponse(false, "Invalid phone or password");

      const officeID = matched[officeIdCol];
      const office = getOfficeById(officesData.rows, officeID);
      const hubName = office ? office[1] : "";

      return jsonResponse(true, "Login success", {
        role: matched[roleCol],
        name: matched[nameCol],
        hubName,
        officeID,
        phone
      });
    }

    // -------- GET OFFICES --------
    if (action === "getOffices") {
      const offices = officesData.rows.map(r => ({
        id: r[0],
        name: r[1],
        number: r[2],
        lat: r[3],
        lng: r[4]
      }));
      return jsonResponse(true, "", { offices });
    }

    // -------- GET OFFICE LOCATION --------
    if (action === "getOfficeLocation") {
      const phone = normalizePhone(getParam(params, 'phone'));
      const emp = employeesData.rows.find(r => normalizePhone(r[phoneCol]) === phone);
      if (!emp) return jsonResponse(false, "Employee not found");

      const officeID = emp[officeIdCol];
      const office = getOfficeById(officesData.rows, officeID);
      if (!office) return jsonResponse(false, "Office not found");

      return jsonResponse(true, "", {
        latitude: office[3],
        longitude: office[4]
      });
    }

    // -------- CHECK-IN / CHECK-OUT --------
    if (action === "Check-In" || action === "Check-Out") {
      const phone = normalizedPhone;
      const shift = getParam(params, 'shift');
      const latitude = parseFloat(params.latitude);
      const longitude = parseFloat(params.longitude);
      const timestamp = new Date();

      if (isNaN(latitude) || isNaN(longitude)) {
        return jsonResponse(false, "Latitude and longitude are required");
      }

      const employeePhones = employeesData.rows.map(r => normalizePhone(r[phoneCol]));
      if (!employeePhones.includes(phone)) {
        return jsonResponse(false, "Phone number not registered");
      }

      let lastRecord = null;
      for (let i = attendanceData.rows.length - 1; i >= 0; i--) {
        const r = attendanceData.rows[i];
        if (normalizePhone(r[1]) === phone && r[3] === shift) {
          lastRecord = r;
          break;
        }
      }

      if (action === "Check-In") {
        if (lastRecord) {
          const lastTime = new Date(lastRecord[0]);
          const diffHours = (timestamp - lastTime) / 3600000;
          if (diffHours < 10) return jsonResponse(false, `Wait ${(10 - diffHours).toFixed(1)} hours to check-in again`);
          if (lastTime.toDateString() === timestamp.toDateString()) return jsonResponse(false, "Already checked-in today for this shift");
        }
      } else {
        if (!lastRecord || lastRecord[4] !== "Check-In") return jsonResponse(false, "No active check-in found");
      }

      let nearestOffice = null;
      let minDist = Infinity;
      officesData.rows.forEach(o => {
        const dist = getDistance(latitude, longitude, o[3], o[4]);
        if (dist < minDist) {
          minDist = dist;
          nearestOffice = o;
        }
      });

      if (!nearestOffice || minDist > 100) {
        return jsonResponse(false, `You are too far from office (${minDist.toFixed(0)} meters)`);
      }

      attendanceSheet.appendRow([
        timestamp,
        phone,
        nearestOffice[0],
        shift,
        action,
        latitude,
        longitude,
        "Active"
      ]);

      return jsonResponse(true, `${action} successful at ${nearestOffice[1]}`, {
        officeName: nearestOffice[1]
      });
    }

    // -------- GET HISTORY --------
    if (action === "getHistory") {
      const phone = normalizePhone(getParam(params, 'phone'));
      const records = attendanceData.rows
        .filter(r => normalizePhone(r[1]) === phone)
        .map(r => ({
          timestamp: new Date(r[0]).toISOString(),
          employeeId: r[1],
          officeId: r[2],
          shift: r[3],
          action: r[4],
          latitude: r[5],
          longitude: r[6],
          status: r[7]
        }));

      return jsonResponse(true, "", { records });
    }

    // -------- GET AGENTS BY OFFICE --------
    if (action === "getAgentsByOffice") {
      const officeID = getParam(params, 'officeID');
      const agents = employeesData.rows
        .filter(r => r[officeIdCol] === officeID && r[roleCol] === Roles.AGENT)
        .map(r => ({ name: r[nameCol], phone: r[phoneCol] }));

      const office = getOfficeById(officesData.rows, officeID);
      const officeName = office ? office[1] : "";

      return jsonResponse(true, "", {
        agents,
        officeName
      });
    }

    // -------- GET ALL EMPLOYEES --------
    if (action === "getAllEmployees") {
      const employees = employeesData.rows.map(r => ({
        name: r[nameCol],
        phone: r[phoneCol],
        role: r[roleCol],
        officeID: r[officeIdCol]
      }));
      return jsonResponse(true, "", { employees });
    }

    return jsonResponse(false, "Invalid action");

  } catch (err) {
    Logger.log(`Unexpected error: ${err.message}`);
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, message: "Server error: " + err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------- HELPERS ----------

function getParam(params, key, required = true) {
  const value = params[key]?.trim();
  if (required && !value) throw new Error(`Missing required parameter: ${key}`);
  return value;
}

function getSheetData(sheet) {
  const data = sheet.getDataRange().getValues();
  return {
    headers: data[0],
    rows: data.slice(1)
  };
}

function getColumnIndexes(headers, columns) {
  return columns.map(col => headers.indexOf(col));
}

function getOfficeById(offices, id) {
  return offices.find(o => o[0] === id);
}

function jsonResponse(success, message, data = {}) {
  return ContentService.createTextOutput(
    JSON.stringify({ success, message, ...data })
  ).setMimeType(ContentService.MimeType.JSON);
}

function normalizePhone(phone) {
  return String(phone).replace(/\D/g, '').trim();
}

function rateLimitExceeded(phone) {
  const LOCK_KEY = 'lastRequestTimestamp_' + phone;
  const RATE_LIMIT_MS = 5000;
  const userProperties = PropertiesService.getScriptProperties();
  const lastRequest = userProperties.getProperty(LOCK_KEY);
  const now = Date.now();

  if (lastRequest && (now - Number(lastRequest) < RATE_LIMIT_MS)) {
    return true;
  }

  userProperties.setProperty(LOCK_KEY, now.toString());
  return false;
}

function getDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const toRad = x => x * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLon / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}
