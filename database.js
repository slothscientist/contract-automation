const Database = require('better-sqlite3');
const path = require('path');

const db = new Database(path.join(__dirname, 'contracts.db'));

// Initialize database
db.exec(`
  CREATE TABLE IF NOT EXISTS contracts (
    id TEXT PRIMARY KEY,
    client_name TEXT NOT NULL,
    client_email TEXT NOT NULL,
    effective_date TEXT,
    term_end_date TEXT,
    payment_1 TEXT,
    payment_2 TEXT,
    launch_date TEXT,
    services_title TEXT,
    services_description TEXT,
    includes_list TEXT,
    what_not_included TEXT,
    project_start_date TEXT,
    project_phases TEXT,
    customer_responsibilities TEXT,
    feedback_policy TEXT,
    completion_policy TEXT,
    cancellation_policy TEXT,
    refund_policy TEXT,
    initials_data TEXT,
    status TEXT DEFAULT 'pending',
    signature_data TEXT,
    signer_name TEXT,
    signed_at DATETIME,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`);

// Migrate existing tables to add new columns
const addColumnIfNotExists = (tableName, columnName, columnType) => {
  try {
    const columns = db.prepare(`PRAGMA table_info(${tableName})`).all();
    const columnExists = columns.some(col => col.name === columnName);
    if (!columnExists) {
      db.exec(`ALTER TABLE ${tableName} ADD COLUMN ${columnName} ${columnType}`);
    }
  } catch (e) {
    // Column might already exist, ignore
  }
};

// Add new Exhibit A columns if they don't exist
addColumnIfNotExists('contracts', 'what_not_included', 'TEXT');
addColumnIfNotExists('contracts', 'project_start_date', 'TEXT');
addColumnIfNotExists('contracts', 'project_phases', 'TEXT');
addColumnIfNotExists('contracts', 'customer_responsibilities', 'TEXT');
addColumnIfNotExists('contracts', 'feedback_policy', 'TEXT');
addColumnIfNotExists('contracts', 'completion_policy', 'TEXT');
addColumnIfNotExists('contracts', 'cancellation_policy', 'TEXT');
addColumnIfNotExists('contracts', 'refund_policy', 'TEXT');
addColumnIfNotExists('contracts', 'initials_data', 'TEXT');

module.exports = {
  createContract: (contract) => {
    const stmt = db.prepare(`
      INSERT INTO contracts (id, client_name, client_email, effective_date, term_end_date,
        payment_1, payment_2, launch_date, services_title, services_description, includes_list,
        what_not_included, project_start_date, project_phases, customer_responsibilities,
        feedback_policy, completion_policy, cancellation_policy, refund_policy)
      VALUES (@id, @client_name, @client_email, @effective_date, @term_end_date,
        @payment_1, @payment_2, @launch_date, @services_title, @services_description, @includes_list,
        @what_not_included, @project_start_date, @project_phases, @customer_responsibilities,
        @feedback_policy, @completion_policy, @cancellation_policy, @refund_policy)
    `);
    return stmt.run(contract);
  },

  getContract: (id) => {
    return db.prepare('SELECT * FROM contracts WHERE id = ?').get(id);
  },

  signContract: (id, signatureData, signerName, initialsData) => {
    const stmt = db.prepare(`
      UPDATE contracts SET status = 'signed', signature_data = ?, signer_name = ?,
        initials_data = ?, signed_at = datetime('now')
      WHERE id = ?
    `);
    return stmt.run(signatureData, signerName, initialsData, id);
  },

  getAllContracts: () => {
    return db.prepare('SELECT id, client_name, client_email, status, created_at, signed_at FROM contracts ORDER BY created_at DESC').all();
  }
};
