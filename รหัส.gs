/**
 * =========================================================
 * CONFIG
 * =========================================================
 */

const SHEET_ID = 'YOUR_SPREADSHEET_ID(เอาไว้ใส่idsheetจริง)';

const QUEUES_SHEET = 'Queues';
const RESERVATIONS_SHEET = 'Reservations';
const SETTINGS_SHEET = 'Settings';

/**
 * ADMIN
 */
const ADMIN_ID = 'XXXXX';
const ADMIN_PASSWORD = 'XXXXX';


/**
 * =========================================================
 * WEB APP
 * =========================================================
 */

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Wuthering Waves Queue')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * =========================================================
 * SHEET MENU
 * =========================================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ จัดการระบบคิว')
    .addItem('🗑️ ล้างคิวทั้งหมด', 'clearAllQueues')
    .addSeparator()
    .addItem('✅ เปิดรับคิว', 'openStore')
    .addItem('⛔ ปิดรับคิวชั่วคราว', 'closeStore')
    .addSeparator()
    .addItem('🔽 ตั้ง Drop-down ระบบจอง', 'setupReservationDropdowns')
    .addToUi();
}


/**
 * =========================================================
 * SHEET HELPERS
 * =========================================================
 */

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SHEET_ID);
}


function getQueueSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(QUEUES_SHEET);

  if (!sheet) {
    throw new Error('ไม่พบชีตชื่อ "' + QUEUES_SHEET + '"');
  }

  return sheet;
}


function getReservationSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(RESERVATIONS_SHEET);

  if (!sheet) {
    throw new Error('ไม่พบชีตชื่อ "' + RESERVATIONS_SHEET + '"');
  }

  return sheet;
}


function getSettingsSheet_() {
  const ss = getSpreadsheet_();

  let sheet = ss.getSheetByName(SETTINGS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);

    sheet.getRange('A1').setValue('สถานะร้าน');
    sheet.getRange('B1').setValue('เปิด');
  }

  return sheet;
}


/**
 * =========================================================
 * STORE STATUS
 * =========================================================
 */

function getStoreStatus() {
  try {
    const val = String(
      getSettingsSheet_().getRange('B1').getValue() || 'เปิด'
    ).trim();

    return val === 'ปิด' ? 'ปิด' : 'เปิด';

  } catch (err) {
    return 'เปิด';
  }
}


function openStore() {
  getSettingsSheet_().getRange('B1').setValue('เปิด');
}


function closeStore() {
  getSettingsSheet_().getRange('B1').setValue('ปิด');
}


/**
 * =========================================================
 * PUBLIC DATA
 * =========================================================
 */

function getData() {
  return {
    queues: getQueues(),
    reservations: getReservations(),
    storeStatus: getStoreStatus()
  };
}


/**
 * =========================================================
 * QUEUES
 *
 * A QueueID
 * B Timestamp
 * C CustomerName
 * D Endgame
 * E DonateAmount
 * F SlipUrl
 * G Note
 * H Status
 * I AdminComment
 * J ContactInfo
 * =========================================================
 */

function getQueues() {
  const sheet = getQueueSheet_();
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return [];
  }

  return data
    .slice(1)
    .filter(row => row[0])
    .map(row => ({
      queueId: String(row[0] || ''),
      timestamp: row[1] instanceof Date
        ? row[1].toLocaleString('th-TH')
        : String(row[1] || ''),
      customerName: String(row[2] || ''),
      endgame: String(row[3] || ''),
      donateAmount: Number(row[4]) || 0,
      slipUrl: String(row[5] || ''),
      note: String(row[6] || ''),
      status: String(row[7] || 'รอเล่น'),
      adminComment: String(row[8] || ''),
      contactInfo: String(row[9] || '')
    }));
}


/**
 * =========================================================
 * RESERVATIONS
 *
 * A ReservationID
 * B Timestamp
 * C CustomerName
 * D Endgame
 * E DonateAmount
 * F SlipUrl
 * G Note
 * H Status
 * I AdminComment
 * J ContactInfo
 * K ApprovedAt
 * L MoveToQueue
 * =========================================================
 */

function getReservations() {
  const sheet = getReservationSheet_();
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return [];
  }

  return data
    .slice(1)
    .filter(row => row[0])
    .map(row => ({
      reservationId: String(row[0] || ''),
      timestamp: row[1] instanceof Date
        ? row[1].toLocaleString('th-TH')
        : String(row[1] || ''),
      customerName: String(row[2] || ''),
      endgame: String(row[3] || ''),
      donateAmount: Number(row[4]) || 0,
      slipUrl: String(row[5] || ''),
      note: String(row[6] || ''),
      status: String(row[7] || 'รออนุมัติ'),
      adminComment: String(row[8] || ''),
      contactInfo: String(row[9] || ''),
      approvedAt: row[10] instanceof Date
        ? row[10].toLocaleString('th-TH')
        : String(row[10] || ''),
      moveToQueue: String(row[11] || 'ยังไม่เข้า')
    }));
}


/**
 * =========================================================
 * NEXT ID
 * =========================================================
 */

function getNextQueueId_(sheet) {
  const data = sheet.getDataRange().getValues();

  let max = 0;

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '');
    const match = id.match(/^Q-(\d+)$/);

    if (match) {
      const num = parseInt(match[1], 10);

      if (num > max) {
        max = num;
      }
    }
  }

  return 'Q-' + String(max + 1).padStart(3, '0');
}


function getNextReservationId_(sheet) {
  const data = sheet.getDataRange().getValues();

  let max = 0;

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '');
    const match = id.match(/^R-(\d+)$/);

    if (match) {
      const num = parseInt(match[1], 10);

      if (num > max) {
        max = num;
      }
    }
  }

  return 'R-' + String(max + 1).padStart(3, '0');
}


/**
 * =========================================================
 * NORMAL QUEUE
 * =========================================================
 */

function submitQueueData(payload) {
  try {

    if (!payload || !payload.customerName) {
      return {
        success: false,
        error: 'กรุณากรอกชื่อ'
      };
    }

    if (!Array.isArray(payload.endgame) || payload.endgame.length === 0) {
      return {
        success: false,
        error: 'กรุณาเลือกโหมด Endgame อย่างน้อย 1 รายการ'
      };
    }

    if (getStoreStatus() === 'ปิด') {
      return {
        success: false,
        error: 'ขณะนี้ร้านปิดรับคิวชั่วคราว'
      };
    }

    const sheet = getQueueSheet_();
    const queueId = getNextQueueId_(sheet);

    const endgame = payload.endgame.join(', ');

    sheet.appendRow([
      queueId,
      new Date(),
      payload.customerName,
      endgame,
      0,
      '',
      payload.note || '',
      'รอเล่น',
      '',
      payload.contactInfo || ''
    ]);

    return {
      success: true,
      queueId: queueId
    };

  } catch (err) {

    return {
      success: false,
      error: err.message
    };

  }
}


/**
 * =========================================================
 * RESERVATION
 * =========================================================
 */

function submitReservation(payload) {
  try {

    if (!payload || !payload.customerName) {
      return {
        success: false,
        error: 'กรุณากรอกชื่อ'
      };
    }

    if (!Array.isArray(payload.endgame) || payload.endgame.length === 0) {
      return {
        success: false,
        error: 'กรุณาเลือกโหมด Endgame อย่างน้อย 1 รายการ'
      };
    }

    const sheet = getReservationSheet_();
    const reservationId = getNextReservationId_(sheet);

    const endgame = payload.endgame.join(', ');

    sheet.appendRow([
      reservationId,
      new Date(),
      payload.customerName,
      endgame,
      0,
      '',
      payload.note || '',
      'รออนุมัติ',
      '',
      payload.contactInfo || '',
      '',
      'ยังไม่เข้า'
    ]);

    return {
      success: true,
      reservationId: reservationId
    };

  } catch (err) {

    return {
      success: false,
      error: err.message
    };

  }
}


/**
 * =========================================================
 * ADMIN LOGIN
 * =========================================================
 */

function adminLogin(user, password) {

  if (
    String(user) === ADMIN_ID &&
    String(password) === ADMIN_PASSWORD
  ) {
    return {
      success: true
    };
  }

  return {
    success: false,
    error: 'Admin ID หรือ Password ไม่ถูกต้อง'
  };
}


/**
 * =========================================================
 * UPDATE RESERVATION
 *
 * Admin แก้:
 * - ชื่อ
 * - Endgame
 * - เงิน
 * - Discord
 * - Note
 * - AdminComment
 * =========================================================
 */

function updateReservationAdmin(payload) {
  try {

    if (!payload || !payload.reservationId) {
      return {
        success: false,
        error: 'ไม่พบเลขจอง'
      };
    }

    const sheet = getReservationSheet_();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {

      if (
        String(data[i][0]) ===
        String(payload.reservationId)
      ) {

        const row = i + 1;

        sheet.getRange(row, 3).setValue(
          payload.customerName || ''
        );

        sheet.getRange(row, 4).setValue(
          payload.endgame || ''
        );

        sheet.getRange(row, 5).setValue(
          Number(payload.donateAmount) || 0
        );

        sheet.getRange(row, 7).setValue(
          payload.note || ''
        );

        sheet.getRange(row, 9).setValue(
          payload.adminComment || ''
        );

        sheet.getRange(row, 10).setValue(
          payload.contactInfo || ''
        );

        return {
          success: true
        };
      }
    }

    return {
      success: false,
      error: 'ไม่พบเลขจอง ' + payload.reservationId
    };

  } catch (err) {

    return {
      success: false,
      error: err.message
    };

  }
}


/**
 * =========================================================
 * APPROVE RESERVATION
 * =========================================================
 */

function approveReservation(reservationId) {
  try {

    const sheet = getReservationSheet_();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {

      if (
        String(data[i][0]) ===
        String(reservationId)
      ) {

        const row = i + 1;

        sheet.getRange(row, 8).setValue('อนุมัติ');
        sheet.getRange(row, 11).setValue(new Date());

        return {
          success: true
        };
      }
    }

    return {
      success: false,
      error: 'ไม่พบเลขจอง'
    };

  } catch (err) {

    return {
      success: false,
      error: err.message
    };

  }
}


/**
 * =========================================================
 * REJECT RESERVATION
 * =========================================================
 */

function rejectReservation(reservationId, comment) {
  try {

    const sheet = getReservationSheet_();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {

      if (
        String(data[i][0]) ===
        String(reservationId)
      ) {

        const row = i + 1;

        sheet.getRange(row, 8).setValue('ปฏิเสธ');

        sheet.getRange(row, 9).setValue(
          comment || ''
        );

        return {
          success: true
        };
      }
    }

    return {
      success: false,
      error: 'ไม่พบเลขจอง'
    };

  } catch (err) {

    return {
      success: false,
      error: err.message
    };

  }
}


/**
 * =========================================================
 * MOVE ONE RESERVATION TO QUEUE
 * =========================================================
 */

function moveReservationToQueue(reservationId) {
  try {

    const reservationSheet = getReservationSheet_();
    const queueSheet = getQueueSheet_();

    const data =
      reservationSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {

      if (
        String(data[i][0]) ===
        String(reservationId)
      ) {

        const row = i + 1;

        const status = String(data[i][7] || '');
        const moved = String(data[i][11] || '');

        if (status !== 'อนุมัติ') {
          return {
            success: false,
            error: 'ต้องอนุมัติการจองก่อน'
          };
        }

        if (moved === 'นำเข้าคิวแล้ว') {
          return {
            success: false,
            error: 'รายการนี้เข้าคิวหลักแล้ว'
          };
        }

        const queueId =
          getNextQueueId_(queueSheet);

        queueSheet.appendRow([
          queueId,
          new Date(),
          data[i][2],
          data[i][3],
          Number(data[i][4]) || 0,
          data[i][5] || '',
          data[i][6] || '',
          'รอเล่น',
          data[i][8] || '',
          data[i][9] || ''
        ]);

        reservationSheet
          .getRange(row, 12)
          .setValue('นำเข้าคิวแล้ว');

        return {
          success: true,
          queueId: queueId
        };
      }
    }

    return {
      success: false,
      error: 'ไม่พบเลขจอง'
    };

  } catch (err) {

    return {
      success: false,
      error: err.message
    };

  }
}


/**
 * =========================================================
 * MOVE ALL APPROVED RESERVATIONS
 * =========================================================
 */

function moveAllApprovedReservations() {
  try {

    const reservationSheet = getReservationSheet_();
    const queueSheet = getQueueSheet_();

    const data =
      reservationSheet.getDataRange().getValues();

    let movedCount = 0;
    let createdQueueIds = [];

    for (let i = 1; i < data.length; i++) {

      const reservationId =
        String(data[i][0] || '');

      const status =
        String(data[i][7] || '');

      const moved =
        String(data[i][11] || '');

      if (
        !reservationId ||
        status !== 'อนุมัติ' ||
        moved === 'นำเข้าคิวแล้ว'
      ) {
        continue;
      }

      const queueId =
        getNextQueueId_(queueSheet);

      queueSheet.appendRow([
        queueId,
        new Date(),
        data[i][2],
        data[i][3],
        Number(data[i][4]) || 0,
        data[i][5] || '',
        data[i][6] || '',
        'รอเล่น',
        data[i][8] || '',
        data[i][9] || ''
      ]);

      reservationSheet
        .getRange(i + 1, 12)
        .setValue('นำเข้าคิวแล้ว');

      movedCount++;
      createdQueueIds.push(queueId);
    }

    return {
      success: true,
      movedCount: movedCount,
      queueIds: createdQueueIds
    };

  } catch (err) {

    return {
      success: false,
      error: err.message
    };

  }
}


/**
 * =========================================================
 * UPDATE QUEUE
 * =========================================================
 */

function updateQueueAdmin(payload) {
  try {

    if (!payload || !payload.queueId) {
      return {
        success: false,
        error: 'ไม่พบเลขคิว'
      };
    }

    const sheet = getQueueSheet_();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {

      if (
        String(data[i][0]) ===
        String(payload.queueId)
      ) {

        const row = i + 1;

        sheet.getRange(row, 3).setValue(
          payload.customerName || ''
        );

        sheet.getRange(row, 4).setValue(
          payload.endgame || ''
        );

        sheet.getRange(row, 5).setValue(
          Number(payload.donateAmount) || 0
        );

        sheet.getRange(row, 7).setValue(
          payload.note || ''
        );

        sheet.getRange(row, 8).setValue(
          payload.status || 'รอเล่น'
        );

        sheet.getRange(row, 9).setValue(
          payload.adminComment || ''
        );

        sheet.getRange(row, 10).setValue(
          payload.contactInfo || ''
        );

        return {
          success: true
        };
      }
    }

    return {
      success: false,
      error: 'ไม่พบเลขคิว ' + payload.queueId
    };

  } catch (err) {

    return {
      success: false,
      error: err.message
    };

  }
}


/**
 * =========================================================
 * DELETE QUEUE
 * =========================================================
 */

function deleteQueueAdmin(queueId) {
  try {

    const sheet = getQueueSheet_();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {

      if (
        String(data[i][0]) ===
        String(queueId)
      ) {

        sheet.deleteRow(i + 1);

        return {
          success: true
        };
      }
    }

    return {
      success: false,
      error: 'ไม่พบเลขคิว'
    };

  } catch (err) {

    return {
      success: false,
      error: err.message
    };

  }
}


/**
 * =========================================================
 * DROPDOWN SETUP
 * =========================================================
 */

function setupReservationDropdowns() {

  const sheet = getReservationSheet_();

  const statusRule =
    SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'รออนุมัติ',
        'อนุมัติ',
        'ปฏิเสธ'
      ], true)
      .setAllowInvalid(false)
      .build();

  const moveRule =
    SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'ยังไม่เข้า',
        'นำเข้าคิวแล้ว'
      ], true)
      .setAllowInvalid(false)
      .build();

  sheet
    .getRange('H2:H1000')
    .setDataValidation(statusRule);

  sheet
    .getRange('L2:L1000')
    .setDataValidation(moveRule);

  return 'ตั้งค่า Drop-down เรียบร้อยแล้ว';
}


/**
 * =========================================================
 * CLEAR QUEUES
 * =========================================================
 */

function clearAllQueues() {

  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    'ยืนยันการล้างคิว',
    'ต้องการลบข้อมูลคิวทั้งหมดใช่หรือไม่?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  const sheet = getQueueSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  ui.alert('ล้างคิวทั้งหมดเรียบร้อยแล้ว');
}


/**
 * =========================================================
 * TEST
 * =========================================================
 */

function testGetQueues() {
  Logger.log(
    JSON.stringify(
      getData(),
      null,
      2
    )
  );
}