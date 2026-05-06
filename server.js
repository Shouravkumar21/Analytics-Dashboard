const express = require('express');
const cors = require('cors');
const multer = require('multer');
const ExcelJS = require('exceljs');
const csv = require('csv-parser');
const fs = require('fs');
const db = require('./db');
require('dotenv').config();

const app = express();
const port = process.env.PORT || 5001;

app.use(cors());
app.use(express.json());
app.use((req, res, next) => {
  console.log(`${req.method} ${req.url}`);
  next();
});

// Test DB Connection
db.query('SELECT NOW()')
  .then(() => console.log('Database connection verified on startup'))
  .catch(err => console.error('Database connection failed on startup:', err));

const upload = multer({ dest: 'uploads/' });

// Root route
app.get('/', (req, res) => {
  res.send('Product Analytics API is running');
});

// Import API
app.post('/api/import', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  const filePath = req.file.path;
  const fileName = req.file.originalname;
  const results = [];

  try {
    if (fileName.endsWith('.csv')) {
      const stream = fs.createReadStream(filePath).pipe(csv());
      for await (const data of stream) {
        results.push(data);
      }
      await processData(results);
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
      res.json({ message: 'CSV data imported successfully' });
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1);
      
      let headers = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          row.eachCell((cell, colNumber) => {
            headers[colNumber] = cell.value ? cell.value.toString().trim() : '';
          });
        } else {
          let rowData = {};
          row.eachCell((cell, colNumber) => {
            if (headers[colNumber]) {
              rowData[headers[colNumber]] = cell.value;
            }
          });
          results.push(rowData);
        }
      });
      await processData(results);
      fs.unlinkSync(filePath);
      res.json({ message: 'Excel data imported successfully' });
    } else {
      res.status(400).json({ error: 'Invalid file format' });
    }
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Internal server error during import' });
  }
});

const format = require('pg-format');

async function processData(data) {
  if (!data || data.length === 0) return;

  const cleanFloat = (val) => {
    if (!val) return 0;
    const str = String(val).replace(/[^0-9.]/g, '');
    return parseFloat(str) || 0;
  };

  const values = data.map(item => {
    // Explicit mapping for Amazon dataset shown in screenshot
    const rawName = item['product_name'] || item['Product Name'] || item['product'] || 'Unknown Product';
    const rawCat = item['category'] || item['Category'] || 'Uncategorized';
    
    // The screenshot shows category is a long string separated by '|'. 
    // Let's take just the first part so it looks better on charts.
    const cleanCat = String(rawCat).split('|')[0] || 'Uncategorized';
    
    const rat = cleanFloat(item['rating']) || cleanFloat(item['Rating']) || 0;
    const rev = cleanFloat(item['rating_count']) || cleanFloat(item['Reviews']) || 0;
    const disc = cleanFloat(item['discount_percentage']) || cleanFloat(item['Discount']) || 0;
    const prc = cleanFloat(item['discounted_price']) || cleanFloat(item['Price']) || 0;

    return [rawName, cleanCat, rat, rev, disc, prc];
  });

  const query = format(
    'INSERT INTO products (product_name, category, rating, review_count, discount, price) VALUES %L',
    values
  );

  try {
    await db.query(query);
    console.log(`Bulk inserted ${data.length} products`);
  } catch (error) {
    console.error('Error during bulk insertion:', error);
    throw error;
  }
}

// Get Products with Search, Filter, Pagination
app.get('/api/products', async (req, res) => {
  const { page = 1, limit = 10, search = '', category = '', minReview = 0 } = req.query;
  const offset = (page - 1) * limit;

  let query = 'SELECT * FROM products WHERE 1=1';
  const params = [];
  let paramIdx = 1;

  if (search) {
    query += ` AND product_name ILIKE $${paramIdx++}`;
    params.push(`%${search}%`);
  }

  if (category) {
    query += ` AND category = $${paramIdx++}`;
    params.push(category);
  }

  if (minReview) {
    query += ` AND review_count >= $${paramIdx++}`;
    params.push(minReview);
  }

  // Count total for pagination
  const countQuery = query.replace('SELECT *', 'SELECT COUNT(*)');
  const totalRes = await db.query(countQuery, params);
  const total = parseInt(totalRes.rows[0].count);

  query += ` ORDER BY id DESC LIMIT $${paramIdx++} OFFSET $${paramIdx++}`;
  params.push(limit, offset);

  try {
    const productsRes = await db.query(query, params);
    res.json({
      data: productsRes.rows,
      total,
      page: parseInt(page),
      limit: parseInt(limit),
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Database error' });
  }
});

// Stats API for Charts
app.get('/api/stats', async (req, res) => {
  try {
    // Products per Category (Limit to top 10 so chart is readable)
    const productsPerCategory = await db.query(
      'SELECT category, COUNT(*) as count FROM products GROUP BY category ORDER BY count DESC LIMIT 10'
    );

    // Top Reviewed Products (Group by name to avoid duplicate bars)
    const topReviewed = await db.query(
      'SELECT product_name, MAX(review_count) as review_count FROM products GROUP BY product_name ORDER BY review_count DESC LIMIT 5'
    );

    // Discount Distribution (Group into bins of 10 for a proper histogram)
    const discountDist = await db.query(
      `SELECT 
         CASE 
           WHEN discount < 10 THEN '0-9%'
           WHEN discount < 20 THEN '10-19%'
           WHEN discount < 30 THEN '20-29%'
           WHEN discount < 40 THEN '30-39%'
           WHEN discount < 50 THEN '40-49%'
           WHEN discount < 60 THEN '50-59%'
           WHEN discount < 70 THEN '60-69%'
           ELSE '70%+' 
         END as discount_bin,
         COUNT(*) as count 
       FROM products 
       GROUP BY discount_bin 
       ORDER BY discount_bin`
    );

    // Category-wise Average Rating (Only for the top 10 most popular categories)
    const avgRating = await db.query(
      'SELECT category, AVG(rating) as average_rating FROM products GROUP BY category ORDER BY COUNT(*) DESC LIMIT 10'
    );

    res.json({
      productsPerCategory: productsPerCategory.rows,
      topReviewed: topReviewed.rows,
      discountDist: discountDist.rows.map(row => ({ discount: row.discount_bin, count: parseInt(row.count) })),
      avgRating: avgRating.rows,
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Database error fetching stats' });
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
