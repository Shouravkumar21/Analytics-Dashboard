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
      
      let headers = {};
      let headerRowFound = false;

      worksheet.eachRow((row, rowNumber) => {
        // Try to find the header row (first row with keywords or just the first row if none found)
        if (!headerRowFound) {
          const rowValues = [];
          row.eachCell((cell) => rowValues.push(String(cell.value || '').toLowerCase()));
          
          const hasKeywords = rowValues.some(v => 
            v.includes('name') || v.includes('product') || v.includes('price') || v.includes('category') || v.includes('rating')
          );

          if (hasKeywords || rowNumber > 10) { // Default to first row or after 10 rows
            row.eachCell((cell, colNumber) => {
              if (cell.value) {
                const cleanHeader = String(cell.value).replace(/[^\x20-\x7E]/g, '').toLowerCase().trim();
                headers[colNumber] = cleanHeader;
              }
            });
            headerRowFound = true;
            return;
          }
        }

        if (headerRowFound) {
          let rowData = {};
          let hasData = false;
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const header = headers[colNumber];
            if (header) {
              const val = cell.value && typeof cell.value === 'object' ? cell.value.result || cell.value.text || cell.value.richText : cell.value;
              rowData[header] = val;
              if (val !== null && val !== undefined && val !== '') hasData = true;
            }
          });
          if (hasData && Object.keys(rowData).length > 0) {
            // Avoid adding the header row itself to results
            const isHeaderRow = Object.values(rowData).some(v => String(v).toLowerCase() === 'product name' || String(v).toLowerCase() === 'price');
            if (!isHeaderRow) results.push(rowData);
          }
        }
      });
      await processData(results);
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
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
  if (!data || data.length === 0) {
    console.log('No data to process');
    return;
  }

  const cleanFloat = (val) => {
    if (val === undefined || val === null || val === '') return 0;
    const str = String(val).trim().replace(/[^0-9.]/g, '');
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  };

  const values = data.map((item, index) => {
    const cleanItem = {};
    Object.keys(item).forEach(key => {
      const cleanKey = key.replace(/[^\x20-\x7E]/g, '').toLowerCase().trim();
      if (cleanKey) cleanItem[cleanKey] = item[key];
    });

    if (index === 0) {
      console.log('Debug - First item keys:', Object.keys(cleanItem));
    }

    // Advanced Keyword Mapping
    const findValue = (keywords) => {
      const key = Object.keys(cleanItem).find(k => keywords.some(kw => k.includes(kw)));
      return key ? cleanItem[key] : null;
    };

    const name = findValue(['product_name', 'product name', 'name', 'title', 'product', 'item']) || 'Unknown Product';
    const rawCat = findValue(['category', 'cat', 'genre', 'type', 'main_category']) || 'Uncategorized';
    const cleanCat = String(rawCat).split('|')[0].trim() || 'Uncategorized';
    
    const rat = cleanFloat(findValue(['rating', 'rate', 'stars', 'score', 'actual_rating']));
    const rev = cleanFloat(findValue(['reviews', 'count', 'rating_count', 'no_of_ratings', 'number', 'review_count']));
    const disc = cleanFloat(findValue(['discount', 'off', 'percent', 'pct', 'discount_percent', 'discount_percentage']));
    const prc = cleanFloat(findValue(['price', 'cost', 'amount', 'value', 'mrp', 'selling_price', 'actual_price']));

    return [name, cleanCat, rat, rev, disc, prc];
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

// Clear database
app.delete('/api/products/clear', async (req, res) => {
  try {
    await db.query('DELETE FROM products');
    res.json({ message: 'Database cleared successfully' });
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
