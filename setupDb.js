const { Client } = require('pg');
require('dotenv').config();

async function init() {
  const client = new Client({
    user: process.env.DB_USER,
    host: process.env.DB_HOST,
    database: process.env.DB_NAME,
    password: process.env.DB_PASSWORD,
    port: process.env.DB_PORT,
    ssl: {
      rejectUnauthorized: false,
    },
  });

  try {
    await client.connect();
    console.log('Connected to Supabase');

    const sql = `
      CREATE TABLE IF NOT EXISTS products (
        id SERIAL PRIMARY KEY,
        product_name TEXT NOT NULL,
        category TEXT NOT NULL,
        rating FLOAT DEFAULT 0,
        review_count INT DEFAULT 0,
        discount FLOAT DEFAULT 0,
        price FLOAT DEFAULT 0,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `;
    await client.query(sql);
    console.log('Table "products" created or already exists');
    await client.end();
  } catch (err) {
    console.error('Error initializing database:', err);
  }
}

init();
