/*******************************************************
 * SONIX 2026
 * DASHBOARD VOTING STAND PAMERAN SONIX 2026
 *
 * FILE : Code.gs
 *
 * FITUR:
 * 1. Voting dengan kode unik 6 karakter
 * 2. Satu kode hanya dapat digunakan satu kali
 * 3. Validasi lokasi / radius
 * 4. Login Admin dengan PIN
 * 5. Generate kode voting
 * 6. Dashboard ranking realtime
 * 7. Security brute-force protection
 *
 * SECURITY:
 *
 * Gagal 1
 *     ↓
 * Gagal 2
 *     ↓
 * Gagal 3
 *     ↓
 * BLOKIR 10 MENIT
 *
 * Setelah 10 menit:
 *
 * Gagal 4
 *     ↓
 * Gagal 5
 *     ↓
 * Gagal 6
 *     ↓
 * Gagal 7
 *     ↓
 * BLOKIR 24 JAM
 *
 *
 * RANKING:
 *
 * 1. Vote terbanyak = ranking lebih tinggi
 *
 * 2. Jika jumlah vote sama:
 *    stand yang LEBIH DAHULU MENCAPAI
 *    JUMLAH VOTE TERSEBUT = ranking lebih tinggi
 *
 * 3. Jika timestamp sama:
 *    baris Votes lebih kecil = lebih tinggi
 *
 * 4. Jika masih sama:
 *    urutan Config
 *******************************************************/


/* =====================================================
   KONFIGURASI UTAMA
===================================================== */

const APP_NAME =
  'Dashboard Voting Stand Pameran Sonix 2026';


const SHEETS = {

  CONFIG: 'Config',

  CODES: 'Kode Voting',

  VOTES: 'Votes',

  SETTINGS: 'Pengaturan'

};


const DEFAULT_STANDS =
  Array.from(
    {
      length: 20
    },
    (_, i) => [
      `Stand ${i + 1}`,
      `STAND-${i + 1}`
    ]
  );


const DEFAULT_LAT =
  -5.147665;


const DEFAULT_LNG =
  119.432732;


const DEFAULT_RADIUS =
  100;


/* =====================================================
   SECURITY
===================================================== */

const SECURITY = {

  CODE_LENGTH:
    6,

  /*
   * Setelah gagal ke-3:
   * blokir 10 menit
   */
  BLOCK_10_MINUTES:
    10 * 60 * 1000,

  /*
   * Setelah gagal ke-7:
   * blokir 24 jam
   */
  BLOCK_24_HOURS:
    24 * 60 * 60 * 1000,

  /*
   * Maksimum gagal:
   * 7
   */
  MAX_FAILED_ATTEMPTS:
    7,

  /*
   * Sesi admin:
   * 30 menit
   */
  ADMIN_SESSION_SECONDS:
    1800,

  /*
   * Prefix penyimpanan
   * security state
   */
  STATE_PREFIX:
    'SONIX_SECURITY_'

};


/* =====================================================
   WEB APP
===================================================== */

function doGet(e) {

  const page =
    e &&
    e.parameter &&
    e.parameter.page

      ? String(
          e.parameter.page
        ).toLowerCase()

      : 'vote';


  // Endpoint bridge untuk halaman Vote yang di-host di GitHub Pages.
  // File HTML Apps Script harus bernama "Bridge".
  if (page === 'bridge') {

    return HtmlService
      .createHtmlOutputFromFile('Bridge')
      .setTitle('SONIX Voting Bridge')
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );

  }


  let fileName =
    'Vote';


  if (
    page === 'dashboard'
  ) {

    fileName =
      'Dashboard';

  }


  else if (
    page === 'admin'
  ) {

    fileName =
      'Admin';

  }


  return HtmlService

    .createTemplateFromFile(
      fileName
    )

    .evaluate()

    .setTitle(
      APP_NAME
    )

    .setXFrameOptionsMode(
      HtmlService
        .XFrameOptionsMode
        .ALLOWALL
    );

}


/* =====================================================
   SETUP APP
===================================================== */

function setupApp() {

  const ss =
    SpreadsheetApp
      .getActive();


  const config =
    getOrCreateSheet_(
      ss,
      SHEETS.CONFIG
    );


  const codes =
    getOrCreateSheet_(
      ss,
      SHEETS.CODES
    );


  const votes =
    getOrCreateSheet_(
      ss,
      SHEETS.VOTES
    );


  const settings =
    getOrCreateSheet_(
      ss,
      SHEETS.SETTINGS
    );


  /* ---------------------------------------------------
     CONFIG
  --------------------------------------------------- */

  if (
    config.getLastRow() === 0
  ) {

    config

      .getRange(
        1,
        1,
        1,
        2
      )

      .setValues([
        [
          'Nama Stand',
          'Kode Stand'
        ]
      ]);


    config

      .getRange(
        2,
        1,
        DEFAULT_STANDS.length,
        2
      )

      .setValues(
        DEFAULT_STANDS
      );


    config.setFrozenRows(
      1
    );

  }


  /* ---------------------------------------------------
     KODE VOTING
  --------------------------------------------------- */

  if (
    codes.getLastRow() === 0
  ) {

    codes

      .getRange(
        1,
        1,
        1,
        4
      )

      .setValues([
        [
          'Kode Voting',
          'Status',
          'Waktu Digunakan',
          'Stand Dipilih'
        ]
      ]);


    codes.setFrozenRows(
      1
    );

  }


  /* ---------------------------------------------------
     VOTES
  --------------------------------------------------- */

  if (
    votes.getLastRow() === 0
  ) {

    votes

      .getRange(
        1,
        1,
        1,
        7
      )

      .setValues([
        [
          'Timestamp',
          'Kode Voting',
          'Kode Stand',
          'Nama Stand',
          'Latitude',
          'Longitude',
          'Jarak (meter)'
        ]
      ]);


    votes.setFrozenRows(
      1
    );

  }


  /* ---------------------------------------------------
     PENGATURAN
  --------------------------------------------------- */

  if (
    settings.getLastRow() === 0
  ) {

    settings

      .getRange(
        1,
        1,
        5,
        2
      )

      .setValues([

        [
          'Parameter',
          'Nilai'
        ],

        [
          'Latitude',
          DEFAULT_LAT
        ],

        [
          'Longitude',
          DEFAULT_LNG
        ],

        [
          'Radius Meter',
          DEFAULT_RADIUS
        ],

        [
          'PIN Admin',
          ''
        ]

      ]);

  }

  else {

    settings

      .getRange(
        'A2:A5'
      )

      .setValues([

        ['Latitude'],

        ['Longitude'],

        ['Radius Meter'],

        ['PIN Admin']

      ]);


    if (
      !settings
        .getRange('B2')
        .getValue()
    ) {

      settings
        .getRange('B2')
        .setValue(
          DEFAULT_LAT
        );

    }


    if (
      !settings
        .getRange('B3')
        .getValue()
    ) {

      settings
        .getRange('B3')
        .setValue(
          DEFAULT_LNG
        );

    }


    if (
      !settings
        .getRange('B4')
        .getValue()
    ) {

      settings
        .getRange('B4')
        .setValue(
          DEFAULT_RADIUS
        );

    }

  }


  /* ---------------------------------------------------
     ADMIN PIN
  --------------------------------------------------- */

  const props =
    PropertiesService
      .getScriptProperties();


  let pin =
    props.getProperty(
      'ADMIN_PIN'
    );


  if (!pin) {

    pin =
      randomDigits_(
        6
      );


    props.setProperty(
      'ADMIN_PIN',
      pin
    );

  }


  settings

    .getRange('B5')

    .setValue(
      pin
    )

    .setNumberFormat(
      '@'
    );


  return {

    success:
      true,

    message:
      'Setup selesai.',

    adminPin:
      pin

  };

}


/* =====================================================
   GET STAND LIST
===================================================== */

function getStandList() {

  const sheet =
    SpreadsheetApp

      .getActive()

      .getSheetByName(
        SHEETS.CONFIG
      );


  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {

    throw new Error(
      'Jalankan setupApp terlebih dahulu.'
    );

  }


  return sheet

    .getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      2
    )

    .getValues()

    .filter(
      row => String(row[1] || '').trim() !== ''
    )
    .map(
      row => ({
        name: String(row[0] || '').trim(),
        code: String(row[1] || '').trim().toUpperCase()
      })
    );

}


/* =====================================================
   GET LOCATION SETTINGS
===================================================== */

function getLocationSettings() {

  const sheet =
    SpreadsheetApp

      .getActive()

      .getSheetByName(
        SHEETS.SETTINGS
      );


  if (!sheet) {

    return {

      latitude:
        DEFAULT_LAT,

      longitude:
        DEFAULT_LNG,

      radius:
        DEFAULT_RADIUS

    };

  }


  return {

    latitude:
      Number(
        sheet
          .getRange('B2')
          .getValue()
      ),

    longitude:
      Number(
        sheet
          .getRange('B3')
          .getValue()
      ),

    radius:
      Number(
        sheet
          .getRange('B4')
          .getValue()
      ) ||
      DEFAULT_RADIUS

  };

}


/* =====================================================
   NORMALIZE CODE
===================================================== */

function normalizeCode_(value) {

  return String(
    value || ''
  )

    .trim()

    .toUpperCase()

    .replace(
      /[^A-Z0-9]/g,
      ''
    );

}


/* =====================================================
   HASH STRING
===================================================== */

function hashString_(value) {

  const digest =
    Utilities.computeDigest(

      Utilities.DigestAlgorithm
        .SHA_256,

      String(
        value || ''
      ),

      Utilities.Charset.UTF_8

    );


  return digest

    .map(
      byte => {

        const v =
          byte < 0
            ? byte + 256
            : byte;


        return (
          '0' +
          v.toString(16)
        ).slice(-2);

      }
    )

    .join('');

}


/* =====================================================
   SECURITY KEY
===================================================== */

function getSecurityKey_(
  clientId
) {

  const normalized =
    String(
      clientId || ''
    ).trim();


  if (
    !normalized
  ) {

    return null;

  }


  return (

    SECURITY.STATE_PREFIX +

    hashString_(
      normalized
    )

  );

}


/* =====================================================
   GET SECURITY STATE
===================================================== */

function getSecurityState_(
  clientId
) {

  const key =
    getSecurityKey_(
      clientId
    );


  if (!key) {

    return {

      failedAttempts:
        0,

      blockedUntil:
        0,

      updatedAt:
        Date.now()

    };

  }


  const props =
    PropertiesService
      .getScriptProperties();


  const raw =
    props.getProperty(
      key
    );


  if (!raw) {

    return {

      failedAttempts:
        0,

      blockedUntil:
        0,

      updatedAt:
        Date.now()

    };

  }


  try {

    const state =
      JSON.parse(
        raw
      );


    return {

      failedAttempts:
        Number(
          state.failedAttempts
        ) || 0,

      blockedUntil:
        Number(
          state.blockedUntil
        ) || 0,

      updatedAt:
        Number(
          state.updatedAt
        ) || 0

    };

  }

  catch (
    error
  ) {

    return {

      failedAttempts:
        0,

      blockedUntil:
        0,

      updatedAt:
        Date.now()

    };

  }

}


/* =====================================================
   SAVE SECURITY STATE
===================================================== */

function saveSecurityState_(
  clientId,
  state
) {

  const key =
    getSecurityKey_(
      clientId
    );


  if (!key) {

    return;

  }


  PropertiesService

    .getScriptProperties()

    .setProperty(

      key,

      JSON.stringify({

        failedAttempts:
          Number(
            state.failedAttempts
          ) || 0,

        blockedUntil:
          Number(
            state.blockedUntil
          ) || 0,

        updatedAt:
          Date.now()

      })

    );

}


/* =====================================================
   CLEAR SECURITY STATE
===================================================== */

function clearSecurityState_(
  clientId
) {

  const key =
    getSecurityKey_(
      clientId
    );


  if (!key) {

    return;

  }


  PropertiesService

    .getScriptProperties()

    .deleteProperty(
      key
    );

}


/* =====================================================
   SECURITY STATUS
===================================================== */

function getVoteSecurityStatus(
  clientId
) {

  const state =
    getSecurityState_(
      clientId
    );


  const now =
    Date.now();


  /* ---------------------------------------------------
     BLOKIR MASIH AKTIF
  --------------------------------------------------- */

  if (

    state.blockedUntil &&

    state.blockedUntil >
      now

  ) {

    return {

      blocked:
        true,

      failedAttempts:
        state.failedAttempts,

      blockedUntil:
        state.blockedUntil,

      remainingMs:
        state.blockedUntil -
        now

    };

  }


  /* ---------------------------------------------------
     BLOKIR SUDAH SELESAI
  ---------------------------------------------------
  
     PENTING:
     failedAttempts TIDAK di-reset.

     Contoh:

     gagal 1
     gagal 2
     gagal 3
     ↓
     blokir 10 menit

     setelah 10 menit:

     failedAttempts = 3

     percobaan berikutnya =
     gagal ke-4
  --------------------------------------------------- */

  if (

    state.blockedUntil &&

    state.blockedUntil <=
      now

  ) {

    state.blockedUntil =
      0;


    saveSecurityState_(
      clientId,
      state
    );

  }


  return {

    blocked:
      false,

    failedAttempts:
      state.failedAttempts,

    blockedUntil:
      0,

    remainingMs:
      0

  };

}


/* =====================================================
   VALIDATE VOTING CODE
===================================================== */

function validateVotingCode(
  code,
  clientId
) {

  code =
    normalizeCode_(
      code
    );


  /* ---------------------------------------------------
     PANJANG KODE
  --------------------------------------------------- */

  if (

    code.length !==
    SECURITY.CODE_LENGTH

  ) {

    return {

      valid:
        false,

      blocked:
        false,

      message:
        'Kode voting harus tepat 6 karakter.'

    };

  }


  /* ---------------------------------------------------
     CEK BLOKIR
  --------------------------------------------------- */

  const security =
    getVoteSecurityStatus(
      clientId
    );


  if (
    security.blocked
  ) {

    return {

      valid:
        false,

      blocked:
        true,

      failedAttempts:
        security.failedAttempts,

      blockedUntil:
        security.blockedUntil,

      remainingMs:
        security.remainingMs,

      message:

        security.failedAttempts >= 7

          ? '7 percobaan gagal. Akses voting diblokir selama 24 jam.'

          : '3 percobaan gagal. Akses voting diblokir selama 10 menit.'

    };

  }


  /* ---------------------------------------------------
     SHEET KODE
  --------------------------------------------------- */

  const sheet =
    SpreadsheetApp

      .getActive()

      .getSheetByName(
        SHEETS.CODES
      );


  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {

    return {

      valid:
        false,

      blocked:
        false,

      message:
        'Belum ada kode voting.'

    };

  }


  const data =
    sheet

      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        4
      )

      .getValues();


  const row =
    data.find(
      r =>
        normalizeCode_(
          r[0]
        ) === code
    );


  /* ---------------------------------------------------
     KODE TIDAK DITEMUKAN
  --------------------------------------------------- */

  if (!row) {

    const lock =
      LockService
        .getScriptLock();


    lock.waitLock(
      10000
    );


    try {

      const state =
        getSecurityState_(
          clientId
        );


      const now =
        Date.now();


      /* -----------------------------------------------
         CEK BLOKIR ULANG
      ----------------------------------------------- */

      if (

        state.blockedUntil &&

        state.blockedUntil >
          now

      ) {

        return {

          valid:
            false,

          blocked:
            true,

          failedAttempts:
            state.failedAttempts,

          blockedUntil:
            state.blockedUntil,

          remainingMs:
            state.blockedUntil -
            now,

          message:

            state.failedAttempts >= 7

              ? '7 percobaan gagal. Akses voting diblokir selama 24 jam.'

              : '3 percobaan gagal. Akses voting diblokir selama 10 menit.'

        };

      }


      /*
       * Jika blokir sudah selesai,
       * hapus hanya blockedUntil.
       *
       * FAILED ATTEMPTS TETAP.
       */

      if (

        state.blockedUntil &&

        state.blockedUntil <=
          now

      ) {

        state.blockedUntil =
          0;

      }


      /* -----------------------------------------------
         TAMBAH GAGAL
      ----------------------------------------------- */

      state.failedAttempts =

        Math.min(

          SECURITY.MAX_FAILED_ATTEMPTS,

          Number(
            state.failedAttempts || 0
          ) + 1

        );


      /* -----------------------------------------------
         ATUR BLOKIR
      ----------------------------------------------- */

      /*
       * HANYA gagal ke-3
       * yang mendapatkan blokir 10 menit.
       */

      if (
        state.failedAttempts === 3
      ) {

        state.blockedUntil =
          now +
          SECURITY.BLOCK_10_MINUTES;

      }


      /*
       * HANYA gagal ke-7
       * yang mendapatkan blokir 24 jam.
       */

      else if (
        state.failedAttempts === 7
      ) {

        state.blockedUntil =
          now +
          SECURITY.BLOCK_24_HOURS;

      }


      /*
       * Gagal 1,2,4,5,6
       * tidak menyebabkan blokir.
       */

      else {

        state.blockedUntil =
          0;

      }


      saveSecurityState_(
        clientId,
        state
      );


      /* -----------------------------------------------
         GAGAL KE-7
      ----------------------------------------------- */

      if (
        state.failedAttempts === 7
      ) {

        return {

          valid:
            false,

          blocked:
            true,

          failedAttempts:
            7,

          blockedUntil:
            state.blockedUntil,

          remainingMs:
            state.blockedUntil -
            now,

          message:
            '7 percobaan gagal. Akses voting diblokir selama 24 jam.'

        };

      }


      /* -----------------------------------------------
         GAGAL KE-3
      ----------------------------------------------- */

      if (
        state.failedAttempts === 3
      ) {

        return {

          valid:
            false,

          blocked:
            true,

          failedAttempts:
            3,

          blockedUntil:
            state.blockedUntil,

          remainingMs:
            state.blockedUntil -
            now,

          message:
            '3 percobaan gagal. Akses voting diblokir selama 10 menit.'

        };

      }


      /* -----------------------------------------------
         GAGAL 1,2,4,5,6
      ----------------------------------------------- */

      return {

        valid:
          false,

        blocked:
          false,

        failedAttempts:
          state.failedAttempts,

        blockedUntil:
          0,

        remainingMs:
          0,

        message:
          'Kode voting tidak ditemukan.'

      };

    }

    finally {

      lock.releaseLock();

    }

  }


  /* ---------------------------------------------------
     KODE SUDAH DIGUNAKAN
  --------------------------------------------------- */

  if (

    String(
      row[1]
    ).toUpperCase() ===
    'USED'

  ) {

    return {

      valid:
        false,

      blocked:
        false,

      message:
        'Kode voting ini sudah digunakan.'

    };

  }


  /* ---------------------------------------------------
     KODE VALID
  --------------------------------------------------- */

  return {

    valid:
      true,

    blocked:
      false,

    failedAttempts:
      security.failedAttempts,

    blockedUntil:
      0,

    message:
      'Kode voting valid.'

  };

}


/* =====================================================
   SUBMIT VOTE
===================================================== */

function submitVote(
  code,
  standCode,
  latitude,
  longitude,
  accuracy,
  clientId
) {

  code =
    normalizeCode_(
      code
    );


  standCode =
    String(
      standCode || ''
    ).trim();


  /* ---------------------------------------------------
     VALIDASI PANJANG KODE
  --------------------------------------------------- */

  if (

    code.length !==
    SECURITY.CODE_LENGTH

  ) {

    return {

      success:
        false,

      message:
        'Kode voting harus tepat 6 karakter.'

    };

  }


  /* ---------------------------------------------------
     CEK SECURITY
  --------------------------------------------------- */

  if (clientId) {

    const security =
      getVoteSecurityStatus(
        clientId
      );


    if (
      security.blocked
    ) {

      return {

        success:
          false,

        blocked:
          true,

        failedAttempts:
          security.failedAttempts,

        blockedUntil:
          security.blockedUntil,

        remainingMs:
          security.remainingMs,

        message:

          security.failedAttempts >= 7

            ? 'Akses voting diblokir selama 24 jam.'

            : 'Akses voting diblokir selama 10 menit.'

      };

    }

  }


  /* ---------------------------------------------------
     LOCK
  --------------------------------------------------- */

  const lock =
    LockService
      .getScriptLock();


  lock.waitLock(
    15000
  );


  try {

    const ss =
      SpreadsheetApp
        .getActive();


    const codeSheet =
      ss.getSheetByName(
        SHEETS.CODES
      );


    const voteSheet =
      ss.getSheetByName(
        SHEETS.VOTES
      );


    if (!codeSheet) {

      return {

        success:
          false,

        message:
          'Sheet Kode Voting tidak ditemukan.'

      };

    }


    if (!voteSheet) {

      return {

        success:
          false,

        message:
          'Sheet Votes tidak ditemukan.'

      };

    }


    /* -------------------------------------------------
       STAND
    ------------------------------------------------- */

    const stand =
      getStandList().find(
        s =>
          s.code ===
          standCode
      );


    if (!stand) {

      return {

        success:
          false,

        message:
          'Stand tidak valid.'

      };

    }


    /* -------------------------------------------------
       KOORDINAT
    ------------------------------------------------- */

    const lat =
      Number(
        latitude
      );


    const lng =
      Number(
        longitude
      );


    const acc =
      Number(
        accuracy
      ) || 0;


    if (

      !isFinite(lat) ||

      !isFinite(lng)

    ) {

      return {

        success:
          false,

        message:
          'Lokasi perangkat tidak valid.'

      };

    }


    if (

      lat < -90 ||

      lat > 90 ||

      lng < -180 ||

      lng > 180

    ) {

      return {

        success:
          false,

        message:
          'Koordinat lokasi tidak valid.'

      };

    }


    /* -------------------------------------------------
       VALIDASI LOKASI
    ------------------------------------------------- */

    const location =
      getLocationSettings();


    const distance =
      haversine_(
        lat,
        lng,
        location.latitude,
        location.longitude
      );


    if (
      distance >
      location.radius
    ) {

      return {

        success:
          false,

        message:
          `Anda berada di luar area voting. Jarak sekitar ${Math.round(distance)} meter. Radius voting ${location.radius} meter.`

      };

    }


    /* -------------------------------------------------
       CARI KODE
    ------------------------------------------------- */

    const lastRow =
      codeSheet.getLastRow();


    if (
      lastRow < 2
    ) {

      return {

        success:
          false,

        message:
          'Belum ada kode voting.'

      };

    }


    const data =
      codeSheet

        .getRange(
          2,
          1,
          lastRow - 1,
          4
        )

        .getValues();


    const index =
      data.findIndex(
        row =>
          normalizeCode_(
            row[0]
          ) === code
      );


    if (
      index < 0
    ) {

      return {

        success:
          false,

        message:
          'Kode voting tidak ditemukan.'

      };

    }


    const rowNumber =
      index + 2;


    /* -------------------------------------------------
       STATUS KODE
    ------------------------------------------------- */

    const status =
      String(
        codeSheet
          .getRange(
            rowNumber,
            2
          )
          .getValue()
      ).toUpperCase();


    if (
      status === 'USED'
    ) {

      return {

        success:
          false,

        message:
          'Kode voting ini sudah digunakan.'

      };

    }


    /* -------------------------------------------------
       CATAT WAKTU
    ------------------------------------------------- */

    const now =
      new Date();


    /*
     * Timestamp vote disimpan
     * pada kolom A Votes.
     *
     * Data ini digunakan untuk
     * menentukan urutan pencapaian
     * jumlah vote.
     */

    voteSheet.appendRow([

      now,

      code,

      stand.code,

      stand.name,

      lat,

      lng,

      distance

    ]);


    /* -------------------------------------------------
       TANDAI KODE USED
    ------------------------------------------------- */

    codeSheet

      .getRange(
        rowNumber,
        2,
        1,
        3
      )

      .setValues([

        [
          'USED',

          now,

          stand.name

        ]

      ]);


    /* -------------------------------------------------
       VOTE BERHASIL
    ------------------------------------------------- */

    if (clientId) {

      clearSecurityState_(
        clientId
      );

    }


    return {

      success:
        true,

      message:
        `Vote untuk ${stand.name} berhasil dicatat.`,

      stand: {

        name:
          stand.name,

        code:
          stand.code

      },

      distance:
        Math.round(
          distance
        ),

      accuracy:
        acc

    };

  }

  finally {

    lock.releaseLock();

  }

}


/* =====================================================
   DASHBOARD DATA
===================================================== */

function getDashboardData() {

  const stands = getStandList();

  const voteSheet = SpreadsheetApp
    .getActive()
    .getSheetByName(SHEETS.VOTES);

  /*
   * KATEGORI DITENTUKAN BERDASARKAN NOMOR KODE STAND.
   * STAND-1 s.d. STAND-15 = Penegak / SMA.
   * STAND-16 s.d. STAND-20 = Penggalang / SMP.
   * Semua kode tetap dibaca agar riwayat vote tidak terputus.
   * Stand dengan nama kosong disembunyikan dari dashboard saja.
   */

  const standData = {};

  stands.forEach((stand, index) => {
    // Kategori mengikuti nomor tetap dalam Kode Stand, bukan index array.
    const match = stand.code.match(/(?:STAND[-_ ]?)(\d+)/i);
    const standNumber = match ? Number(match[1]) : null;
    const category = standNumber !== null
      ? (standNumber >= 1 && standNumber <= 15
          ? 'penegak'
          : (standNumber >= 16 && standNumber <= 20 ? 'penggalang' : 'lainnya'))
      : (index < 15 ? 'penegak' : 'penggalang');

    // Gunakan kunci normalisasi agar "Stand 18", "STAND 18",
    // dan "STAND-18" tetap merujuk ke stand yang sama.
    const normalizedStandCode = normalizeCode_(stand.code);

    standData[normalizedStandCode] = {
      code: stand.code,
      name: stand.name,
      active: String(stand.name || '').trim() !== '',
      votes: 0,
      lastVoteTime: null,
      lastVoteRow: Number.MAX_SAFE_INTEGER,
      configOrder: index,
      category: category
    };
  });

  /*
   * Baca data voting. Setiap suara dihitung berdasarkan Kode Stand
   * yang tersimpan pada kolom C sheet Votes.
   */

  if (voteSheet && voteSheet.getLastRow() > 1) {
    const rows = voteSheet
      .getRange(2, 1, voteSheet.getLastRow() - 1, 7)
      .getValues();

    rows.forEach((row, index) => {
      const timestamp = row[0];
      // Kolom C pada Votes berisi Kode Stand. Normalisasi harus sama
      // dengan kunci dari Config agar perbedaan spasi/tanda hubung/huruf
      // besar-kecil tidak membuat vote lama terlewat.
      const standCode = normalizeCode_(row[2]);

      if (!standCode || !standData[standCode]) return;

      const stand = standData[standCode];
      stand.votes++;

      const sheetRow = index + 2;
      let voteTime = null;

      if (timestamp instanceof Date && !isNaN(timestamp.getTime())) {
        voteTime = timestamp.getTime();
      }

      if (voteTime !== null) {
        stand.lastVoteTime = voteTime;
      }

      stand.lastVoteRow = sheetRow;
    });
  }

  /*
   * Fungsi ranking terpisah untuk setiap kategori.
   * Aturan tie-break dipertahankan:
   * 1. Vote terbanyak
   * 2. Lebih dahulu mencapai jumlah vote saat ini
   * 3. Baris Votes lebih kecil
   * 4. Urutan Config
   */

  function sortStandings(list) {
    return list.sort((a, b) => {
      if (a.votes !== b.votes) {
        return b.votes - a.votes;
      }

      if (a.votes === 0 && b.votes === 0) {
        return a.configOrder - b.configOrder;
      }

      if (
        a.lastVoteTime !== null &&
        b.lastVoteTime !== null &&
        a.lastVoteTime !== b.lastVoteTime
      ) {
        return a.lastVoteTime - b.lastVoteTime;
      }

      if (a.lastVoteRow !== b.lastVoteRow) {
        return a.lastVoteRow - b.lastVoteRow;
      }

      return a.configOrder - b.configOrder;
    });
  }

  function publicStand(stand) {
    return {
      code: stand.code,
      name: stand.name,
      votes: stand.votes
    };
  }

  function buildCategory(categoryKey) {
    const categoryStands = Object.values(standData)
      .filter(stand => stand.category === categoryKey && stand.active);

    const ranked = sortStandings(categoryStands);

    const totalVotes = ranked.reduce(
      (total, stand) => total + stand.votes,
      0
    );

    const champion = totalVotes > 0 && ranked.length > 0
      ? publicStand(ranked[0])
      : null;

    return {
      champion: champion,
      stands: ranked.map(publicStand),
      totalVotes: totalVotes
    };
  }

  const penegak = buildCategory('penegak');
  const penggalang = buildCategory('penggalang');

  const totalVotes = penegak.totalVotes + penggalang.totalVotes;

  /*
   * Properti champion dan stands lama tetap disediakan untuk kompatibilitas
   * sementara dengan Dashboard.html versi lama.
   * Dashboard dua kategori harus membaca:
   * data.penegak.champion, data.penegak.stands
   * data.penggalang.champion, data.penggalang.stands
   */

  const allRanked = sortStandings(Object.values(standData));

  return {
    appName: APP_NAME,
    totalVotes: totalVotes,

    penegak: penegak,
    penggalang: penggalang,

    champions: {
      penegak: penegak.champion,
      penggalang: penggalang.champion
    },

    categories: {
      penegak: penegak,
      penggalang: penggalang
    },

    // Legacy fields; jangan dipakai untuk kartu juara kategori.
    champion: totalVotes > 0 && allRanked.length > 0
      ? publicStand(allRanked[0])
      : null,

    stands: allRanked.map(publicStand),

    updatedAt: new Date().toISOString()
  };
}

/* =====================================================
   ADMIN LOGIN
===================================================== */

function adminLogin(
  pin
) {

  const savedPin =
    PropertiesService

      .getScriptProperties()

      .getProperty(
        'ADMIN_PIN'
      );


  if (

    String(
      pin || ''
    ) !==

    String(
      savedPin || ''
    )

  ) {

    return {

      success:
        false,

      message:
        'PIN Admin salah.'

    };

  }


  const token =
    Utilities.getUuid();


  CacheService

    .getScriptCache()

    .put(

      'ADMIN_' +
      token,

      'OK',

      SECURITY
        .ADMIN_SESSION_SECONDS

    );


  return {

    success:
      true,

    token:
      token,

    expiresIn:
      SECURITY
        .ADMIN_SESSION_SECONDS

  };

}


/* =====================================================
   REQUIRE ADMIN
===================================================== */

function requireAdmin_(
  token
) {

  if (!token) {

    throw new Error(
      'Sesi Admin tidak valid.'
    );

  }


  const value =
    CacheService

      .getScriptCache()

      .get(
        'ADMIN_' +
        token
      );


  if (
    value !== 'OK'
  ) {

    throw new Error(
      'Sesi Admin tidak valid atau telah kedaluwarsa.'
    );

  }

}


/* =====================================================
   GENERATE VOTING CODES
===================================================== */

function generateVotingCodes(
  token,
  quantity
) {

  requireAdmin_(
    token
  );


  quantity =
    Math.floor(
      Number(
        quantity
      )
    );


  if (

    quantity < 1 ||

    quantity > 5000

  ) {

    throw new Error(
      'Jumlah kode harus 1-5000.'
    );

  }


  const sheet =
    SpreadsheetApp

      .getActive()

      .getSheetByName(
        SHEETS.CODES
      );


  if (!sheet) {

    throw new Error(
      'Sheet Kode Voting tidak ditemukan.'
    );

  }


  const used =
    new Set();


  if (
    sheet.getLastRow() > 1
  ) {

    sheet

      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        1
      )

      .getValues()

      .forEach(
        row => {

          const code =
            normalizeCode_(
              row[0]
            );


          if (code) {

            used.add(
              code
            );

          }

        }
      );

  }


  const rows = [];


  while (

    rows.length <
    quantity

  ) {

    const code =
      generateCode_();


    if (
      !used.has(code)
    ) {

      used.add(
        code
      );


      rows.push([

        code,

        'UNUSED',

        '',

        ''

      ]);

    }

  }


  sheet

    .getRange(

      sheet.getLastRow() + 1,

      1,

      rows.length,

      4

    )

    .setValues(
      rows
    );


  return {

    success:
      true,

    generated:
      rows.map(
        row =>
          row[0]
      )

  };

}


/* =====================================================
   CODE SUMMARY
===================================================== */

function getCodeSummary(
  token
) {

  requireAdmin_(
    token
  );


  const sheet =
    SpreadsheetApp

      .getActive()

      .getSheetByName(
        SHEETS.CODES
      );


  if (

    !sheet ||

    sheet.getLastRow() < 2

  ) {

    return {

      total:
        0,

      unused:
        0,

      used:
        0

    };

  }


  const data =
    sheet

      .getRange(

        2,

        1,

        sheet.getLastRow() - 1,

        4

      )

      .getValues();


  const used =
    data.filter(
      row =>
        String(
          row[1]
        ).toUpperCase() ===
        'USED'
    ).length;


  return {

    total:
      data.length,

    unused:
      data.length -
      used,

    used:
      used

  };

}


/* =====================================================
   SAVE LOCATION SETTINGS
===================================================== */

function saveLocationSettings(
  token,
  latitude,
  longitude,
  radius
) {

  requireAdmin_(
    token
  );


  latitude =
    Number(
      latitude
    );


  longitude =
    Number(
      longitude
    );


  radius =
    Number(
      radius
    );


  if (

    !isFinite(
      latitude
    ) ||

    !isFinite(
      longitude
    ) ||

    !isFinite(
      radius
    ) ||

    radius <= 0

  ) {

    throw new Error(
      'Lokasi/radius tidak valid.'
    );

  }


  if (

    latitude < -90 ||

    latitude > 90

  ) {

    throw new Error(
      'Latitude tidak valid.'
    );

  }


  if (

    longitude < -180 ||

    longitude > 180

  ) {

    throw new Error(
      'Longitude tidak valid.'
    );

  }


  const sheet =
    SpreadsheetApp

      .getActive()

      .getSheetByName(
        SHEETS.SETTINGS
      );


  sheet

    .getRange(
      'B2:B4'
    )

    .setValues([

      [latitude],

      [longitude],

      [radius]

    ]);


  return {

    success:
      true,

    latitude:
      latitude,

    longitude:
      longitude,

    radius:
      radius

  };

}


/* =====================================================
   GENERATE CODE
===================================================== */

function generateCode_() {

  /*
   * I, O, 0 dan 1
   * tidak digunakan.
   */

  const chars =
    'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';


  let code =
    '';


  for (

    let i = 0;

    i <
    SECURITY.CODE_LENGTH;

    i++

  ) {

    code +=

      chars.charAt(

        Math.floor(

          Math.random() *
          chars.length

        )

      );

  }


  return code;

}


/* =====================================================
   RANDOM DIGITS
===================================================== */

function randomDigits_(
  length
) {

  let result =
    '';


  for (

    let i = 0;

    i < length;

    i++

  ) {

    result +=

      Math.floor(

        Math.random() *
        10

      );

  }


  return result;

}


/* =====================================================
   HAVERSINE
===================================================== */

function haversine_(
  lat1,
  lon1,
  lat2,
  lon2
) {

  const R =
    6371000;


  const toRad =
    Math.PI / 180;


  const dLat =
    (
      lat2 -
      lat1
    ) *
    toRad;


  const dLon =
    (
      lon2 -
      lon1
    ) *
    toRad;


  const a =

    Math.sin(
      dLat / 2
    ) *
    Math.sin(
      dLat / 2
    )

    +

    Math.cos(
      lat1 *
      toRad
    )

    *

    Math.cos(
      lat2 *
      toRad
    )

    *

    Math.sin(
      dLon / 2
    ) *
    Math.sin(
      dLon / 2
    );


  return (

    R *

    2 *

    Math.atan2(

      Math.sqrt(a),

      Math.sqrt(
        1 - a
      )

    )

  );

}


/* =====================================================
   GET OR CREATE SHEET
===================================================== */

function getOrCreateSheet_(
  ss,
  name
) {

  return (

    ss.getSheetByName(
      name
    )

    ||

    ss.insertSheet(
      name
    )

  );

}


/* =====================================================
   RESET ALL VOTING CODES
===================================================== */

function resetAllVotingCodes(
  token
) {

  requireAdmin_(
    token
  );


  const sheet =
    SpreadsheetApp

      .getActive()

      .getSheetByName(
        SHEETS.CODES
      );


  if (

    !sheet ||

    sheet.getLastRow() < 2

  ) {

    return {

      success:
        true,

      message:
        'Tidak ada kode voting untuk di-reset.'

    };

  }


  const count =
    sheet.getLastRow() -
    1;


  /*
   * Bersihkan:
   *
   * Status
   * Waktu
   * Stand
   */

  sheet

    .getRange(
      2,
      2,
      count,
      3
    )

    .clearContent();


  /*
   * Semua kembali UNUSED.
   */

  sheet

    .getRange(
      2,
      2,
      count,
      1
    )

    .setValues(

      Array.from(

        {
          length:
            count
        },

        () => [
          'UNUSED'
        ]

      )

    );


  /*
   * Reset security state.
   */

  clearAllSecurityStates_();


  return {

    success:
      true,

    message:
      'Semua kode voting berhasil di-reset.'

  };

}


/* =====================================================
   RESET CLIENT SECURITY
===================================================== */

function resetClientSecurity(
  token,
  clientId
) {

  requireAdmin_(
    token
  );


  if (!clientId) {

    throw new Error(
      'Client ID tidak boleh kosong.'
    );

  }


  clearSecurityState_(
    clientId
  );


  return {

    success:
      true,

    message:
      'Security state client berhasil di-reset.'

  };

}


/* =====================================================
   CLEAR ALL SECURITY STATES
===================================================== */

function clearAllSecurityStates_() {

  const props =
    PropertiesService
      .getScriptProperties();


  const all =
    props.getProperties();


  const keys =
    Object.keys(
      all
    )

    .filter(
      key =>
        key.indexOf(
          SECURITY.STATE_PREFIX
        ) === 0
    );


  keys.forEach(
    key =>
      props.deleteProperty(
        key
      )
  );

}


/* =====================================================
   SYSTEM INFO
===================================================== */

function getSystemInfo(
  token
) {

  requireAdmin_(
    token
  );


  const location =
    getLocationSettings();


  const codes =
    getCodeSummary(
      token
    );


  return {

    appName:
      APP_NAME,

    location:
      location,

    codes:
      codes,

    timestamp:
      new Date()
        .toISOString()

  };

}