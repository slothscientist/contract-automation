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
    status TEXT DEFAULT 'pending',
    signature_data TEXT,
    signer_name TEXT,
    signed_at DATETIME,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`);

module.exports = {
  createContract: (contract) => {
    const stmt = db.prepare(`
      INSERT INTO contracts (id, client_name, client_email, effective_date, term_end_date,
        payment_1, payment_2, launch_date, services_title, services_description, includes_list)
      VALUES (@id, @client_name, @client_email, @effective_date, @term_end_date,
        @payment_1, @payment_2, @launch_date, @services_title, @services_description, @includes_list)
    `);
    return stmt.run(contract);
  },

  getContract: (id) => {
    return db.prepare('SELECT * FROM contracts WHERE id = ?').get(id);
  },

  signContract: (id, signatureData, signerName) => {
    const stmt = db.prepare(`
      UPDATE contracts SET status = 'signed', signature_data = ?, signer_name = ?, signed_at = datetime('now')
      WHERE id = ?
    `);
    return stmt.run(signatureData, signerName, id);
  },

  getAllContracts: () => {
    return db.prepare('SELECT id, client_name, client_email, status, created_at, signed_at FROM contracts ORDER BY created_at DESC').all();
  }
};
