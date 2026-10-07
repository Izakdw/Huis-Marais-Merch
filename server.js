const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
// Support both nested folders (local) and flat file layout (cloud upload)
// { index: false } prevents accidental public exposure of index.html at root
app.use(express.static(path.join(__dirname, 'public'), { index: false }));
app.use(express.static(__dirname, { index: false }));
app.use('/images/merch', express.static(path.join(__dirname, 'public', 'images', 'merch')));
app.use('/images/merch', express.static(__dirname));
app.use('/images', express.static(path.join(__dirname, 'public', 'images')));
app.use('/images', express.static(__dirname));

const DATA_DIR = fs.existsSync(path.join(__dirname, 'data')) ? path.join(__dirname, 'data') : __dirname;
const STORE_FILE = fs.existsSync(path.join(DATA_DIR, 'store.json')) 
  ? path.join(DATA_DIR, 'store.json') 
  : (fs.existsSync(path.join(__dirname, 'store.json')) ? path.join(__dirname, 'store.json') : path.join(DATA_DIR, 'store.json'));
const INITIAL_DATA_FILE = fs.existsSync(path.join(DATA_DIR, 'initial_data.json'))
  ? path.join(DATA_DIR, 'initial_data.json')
  : path.join(__dirname, 'initial_data.json');

// Ensure data folder exists
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

// DEFAULT STOREFRONT FOR RESIDENTS: Root (/) serves the student catalog (NO profits/costs)
app.get('/', (req, res) => {
  const localCatalog = path.join(__dirname, 'public', 'catalog.html');
  if (fs.existsSync(localCatalog)) return res.sendFile(localCatalog);
  return res.sendFile(path.join(__dirname, 'catalog.html'));
});

// Initialize store if not present
function getStore() {
  if (!fs.existsSync(STORE_FILE)) {
    let initialProducts = [];
    if (fs.existsSync(INITIAL_DATA_FILE)) {
      const initData = JSON.parse(fs.readFileSync(INITIAL_DATA_FILE, 'utf-8'));
      initialProducts = initData.products || [];
    }

    const inventory = {};
    initialProducts.forEach(p => {
      p.sizes.forEach(size => {
        const key = `${p.id}__${size}`;
        inventory[key] = {
          productId: p.id,
          productName: p.name,
          size: size,
          openingQty: 0,
          arrivedQty: 0,
          countedQty: null
        };
      });
    });

    const store = {
      products: initialProducts,
      inventory: inventory,
      sales: [],
      arrivals: []
    };

    fs.writeFileSync(STORE_FILE, JSON.stringify(store, null, 2));
    return store;
  }

  return JSON.parse(fs.readFileSync(STORE_FILE, 'utf-8'));
}

function saveStore(store) {
  const jsonContent = JSON.stringify(store, null, 2);
  fs.writeFileSync(STORE_FILE, jsonContent);

  // Automatic backup safeguard
  try {
    const backupDir = path.join(DATA_DIR, 'backups');
    if (!fs.existsSync(backupDir)) fs.mkdirSync(backupDir, { recursive: true });
    const backupFile = path.join(backupDir, 'store_latest_backup.json');
    fs.writeFileSync(backupFile, jsonContent);
  } catch (e) {
    console.error('Backup error:', e);
  }
}

// Helper: Calculate inventory stats
function calculateMetrics(store) {
  const sales = store.sales || [];
  const inventory = store.inventory || {};
  const products = store.products || [];
  const productMap = new Map(products.map(p => [p.id, p]));

  // Sold count per product-size
  const soldMap = {};
  let totalRevenue = 0;
  let totalCostOfGoodsSold = 0;
  let totalUnitsSold = 0;

  sales.forEach(sale => {
    const key = `${sale.productId}__${sale.size}`;
    soldMap[key] = (soldMap[key] || 0) + sale.qty;
    totalRevenue += Number(sale.amount) || 0;
    const cp = Number(sale.costPrice) || 0;
    totalCostOfGoodsSold += cp * sale.qty;
    totalUnitsSold += sale.qty;
  });

  const totalProfit = totalRevenue - totalCostOfGoodsSold;
  const marginPct = totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : 0;

  // Build stock items
  const stockItems = [];
  let totalStockRemaining = 0;
  let totalStockValuationCost = 0;
  let totalStockValuationRetail = 0;

  Object.keys(inventory).forEach(key => {
    const item = inventory[key];
    const product = productMap.get(item.productId);
    const sold = soldMap[key] || 0;
    const received = (item.openingQty || 0) + (item.arrivedQty || 0);
    const remaining = received - sold;
    const cp = product ? Number(product.costPrice) || 0 : 0;
    const cpVat = product ? Number(product.costPriceVat) || 0 : 0;
    const sp = product ? Number(product.sellingPrice) || 0 : 0;
    const counted = item.countedQty !== null && item.countedQty !== undefined ? Number(item.countedQty) : null;
    const variance = counted !== null ? counted - remaining : null;

    let status = 'In Stock';
    if (remaining <= 0) {
      status = 'Out of Stock';
    } else if (remaining <= 3) {
      status = 'Low Stock';
    }

    totalStockRemaining += remaining;
    totalStockValuationCost += remaining * cp;
    totalStockValuationRetail += remaining * sp;

    stockItems.push({
      key,
      productId: item.productId,
      productName: item.productName,
      category: product ? product.category : 'General',
      size: item.size,
      costPrice: cp,
      costPriceVat: cpVat,
      sellingPrice: sp,
      profitPerUnit: sp - (cpVat || cp),
      openingQty: item.openingQty || 0,
      arrivedQty: item.arrivedQty || 0,
      totalReceived: received,
      soldToDate: sold,
      stockOnHand: remaining,
      countedQty: counted,
      variance: variance,
      status: status
    });
  });

  return {
    summary: {
      totalRevenue,
      totalCostOfGoodsSold,
      totalProfit,
      marginPct,
      totalUnitsSold,
      totalStockRemaining,
      totalStockValuationCost,
      totalStockValuationRetail,
      potentialRemainingProfit: totalStockValuationRetail - totalStockValuationCost
    },
    stockItems
  };
}

// ---------------- API ROUTES ----------------

// GET Full State
app.get('/catalog', (req, res) => {
  const localCatalog = path.join(__dirname, 'public', 'catalog.html');
  if (fs.existsSync(localCatalog)) return res.sendFile(localCatalog);
  return res.sendFile(path.join(__dirname, 'catalog.html'));
});

app.get('/catalog.html', (req, res) => {
  const localCatalog = path.join(__dirname, 'public', 'catalog.html');
  if (fs.existsSync(localCatalog)) return res.sendFile(localCatalog);
  return res.sendFile(path.join(__dirname, 'catalog.html'));
});

// PRIVATE ADMIN PORTAL (Only for Izak to manage sales/stock)
app.get('/admin', (req, res) => {
  const localIndex = path.join(__dirname, 'public', 'index.html');
  if (fs.existsSync(localIndex)) return res.sendFile(localIndex);
  return res.sendFile(path.join(__dirname, 'index.html'));
});

// If anyone visits index.html directly, redirect them to the safe resident catalog
app.get('/index.html', (req, res) => {
  res.redirect('/');
});

// GET Public Catalog Data (NO cost prices, NO profit margins, NO sales logs!)
app.get('/api/public-catalog', (req, res) => {
  try {
    const store = getStore();
    const metrics = calculateMetrics(store);
    
    // Map stock items by product ID
    const stockByProduct = {};
    metrics.stockItems.forEach(item => {
      if (!stockByProduct[item.productId]) {
        stockByProduct[item.productId] = [];
      }
      stockByProduct[item.productId].push({
        size: item.size,
        stockOnHand: item.stockOnHand,
        status: item.status
      });
    });

    const imageMap = {
      '80th-year-pullover': '/image8.png',
      'dress-shirt': '/image11.png',
      '80th-year-hat': '/image12.png',
      '80th-rugby-jersey': '/image13.png',
      'hm-80-years-t-shirt': '/image14.png',
      'white-t-shirt-no-6': '/image16.png',
      'alumni-cap': '/image19.png',
      'first-years-shirt': '/image20.png',
      'vintage-polo-shirt': '/vintage-polo-both.jpeg',
      'rugby-jersey-green-yellow': '/image24.jpeg'
    };

    const publicCatalog = store.products.map(p => {
      const sizes = stockByProduct[p.id] || p.sizes.map(s => ({ size: s, stockOnHand: 0, status: 'Out of Stock' }));
      const totalAvailable = sizes.reduce((sum, s) => sum + Math.max(0, s.stockOnHand), 0);

      return {
        id: p.id,
        name: p.name,
        category: p.category,
        colour: p.colour,
        sellingPrice: p.sellingPrice,
        image: p.image || imageMap[p.id] || null,
        sizes: sizes,
        totalAvailable: totalAvailable,
        inStock: totalAvailable > 0
      };
    });

    res.json({
      residenceName: 'Huis Marais • 80 Jaar Merch',
      logo: '/images/80th-logo.png',
      emblem: '/images/emblem.png',
      items: publicCatalog
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET Full State for Admin Portal
app.get('/api/data', (req, res) => {
  try {
    const store = getStore();
    const metrics = calculateMetrics(store);
    res.json({
      products: store.products,
      sales: store.sales,
      arrivals: store.arrivals,
      summary: metrics.summary,
      stockItems: metrics.stockItems
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// POST Merch Arrival (Stock Intake)
app.post('/api/arrivals', (req, res) => {
  try {
    const { productId, date, sizeBreakdown, notes } = req.body;
    if (!productId || !sizeBreakdown) {
      return res.status(400).json({ error: 'Product ID and sizeBreakdown are required' });
    }

    const store = getStore();
    const product = store.products.find(p => p.id === productId);
    if (!product) {
      return res.status(404).json({ error: 'Product not found' });
    }

    let totalUnits = 0;
    Object.keys(sizeBreakdown).forEach(size => {
      const qty = parseInt(sizeBreakdown[size], 10) || 0;
      if (qty > 0) {
        totalUnits += qty;
        const key = `${productId}__${size}`;
        if (!store.inventory[key]) {
          store.inventory[key] = {
            productId,
            productName: product.name,
            size,
            openingQty: 0,
            arrivedQty: 0,
            countedQty: null
          };
        }
        store.inventory[key].arrivedQty = (store.inventory[key].arrivedQty || 0) + qty;
      }
    });

    const arrivalRecord = {
      id: 'arr_' + Date.now(),
      date: date || new Date().toISOString().split('T')[0],
      productId,
      productName: product.name,
      sizeBreakdown,
      totalUnits,
      notes: notes || ''
    };

    store.arrivals.unshift(arrivalRecord);
    saveStore(store);

    const metrics = calculateMetrics(store);
    res.json({
      success: true,
      message: `Successfully logged arrival of ${totalUnits} units for ${product.name}`,
      arrival: arrivalRecord,
      summary: metrics.summary,
      stockItems: metrics.stockItems
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// DELETE a merch arrival (undo/reverses stock addition)
app.delete('/api/arrivals/:id', (req, res) => {
  try {
    const store = getStore();
    const index = store.arrivals.findIndex(a => a.id === req.params.id);
    if (index === -1) {
      return res.status(404).json({ error: 'Arrival record not found' });
    }

    const arrival = store.arrivals.splice(index, 1)[0];

    // Deduct the quantities that were added by this arrival
    if (arrival.sizeBreakdown) {
      Object.entries(arrival.sizeBreakdown).forEach(([size, qty]) => {
        const key = `${arrival.productId}__${size}`;
        if (store.inventory[key]) {
          store.inventory[key].arrivedQty = Math.max(0, (store.inventory[key].arrivedQty || 0) - (parseInt(qty, 10) || 0));
        }
      });
    }

    saveStore(store);

    const metrics = calculateMetrics(store);
    res.json({
      success: true,
      message: `Deleted arrival for ${arrival.productName} (${arrival.totalUnits} units removed from stock)`,
      summary: metrics.summary,
      stockItems: metrics.stockItems
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// POST Record a Sale (Per Person)
app.post('/api/sales', (req, res) => {
  try {
    const { date, buyer, productId, size, qty, unitPrice, notes } = req.body;
    if (!buyer || !productId || !size || !qty) {
      return res.status(400).json({ error: 'Buyer, product, size, and quantity are required' });
    }

    const quantity = parseInt(qty, 10);
    if (isNaN(quantity) || quantity <= 0) {
      return res.status(400).json({ error: 'Quantity must be a positive integer' });
    }

    const store = getStore();
    const product = store.products.find(p => p.id === productId);
    if (!product) {
      return res.status(404).json({ error: 'Product not found' });
    }

    const invKey = `${productId}__${size}`;
    const invItem = store.inventory[invKey];
    const totalReceived = invItem ? ((invItem.openingQty || 0) + (invItem.arrivedQty || 0)) : 0;
    
    // Check current sales for this item
    const currentSold = store.sales
      .filter(s => s.productId === productId && s.size === size)
      .reduce((sum, s) => sum + s.qty, 0);

    const availableStock = totalReceived - currentSold;
    if (availableStock < quantity) {
      return res.status(400).json({
        error: `Insufficient stock for ${product.name} (${size}). Only ${availableStock} remaining.`
      });
    }

    const price = unitPrice !== undefined && unitPrice !== '' ? Number(unitPrice) : product.sellingPrice;
    const amount = price * quantity;
    const costPrice = product.costPriceVat || product.costPrice || 0;
    const profit = amount - (costPrice * quantity);

    const saleRecord = {
      id: 'sale_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4),
      date: date || new Date().toISOString().split('T')[0],
      buyer: buyer.trim(),
      productId,
      productName: product.name,
      size,
      qty: quantity,
      unitPrice: price,
      amount,
      costPrice,
      profit,
      notes: notes || ''
    };

    store.sales.unshift(saleRecord);
    saveStore(store);

    const metrics = calculateMetrics(store);
    res.json({
      success: true,
      message: `Sale recorded for ${buyer}: ${quantity}x ${product.name} (${size})`,
      sale: saleRecord,
      summary: metrics.summary,
      stockItems: metrics.stockItems
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// DELETE a sale (reverses entry)
app.delete('/api/sales/:id', (req, res) => {
  try {
    const store = getStore();
    const index = store.sales.findIndex(s => s.id === req.params.id);
    if (index === -1) {
      return res.status(404).json({ error: 'Sale record not found' });
    }

    const deletedSale = store.sales.splice(index, 1)[0];
    saveStore(store);

    const metrics = calculateMetrics(store);
    res.json({
      success: true,
      message: `Deleted sale for ${deletedSale.buyer}`,
      summary: metrics.summary,
      stockItems: metrics.stockItems
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// POST Count check (physical stock take)
app.post('/api/stock-take/count', (req, res) => {
  try {
    const { key, countedQty } = req.body;
    const store = getStore();

    if (!store.inventory[key]) {
      return res.status(404).json({ error: 'Stock item not found' });
    }

    store.inventory[key].countedQty = countedQty === '' || countedQty === null ? null : parseInt(countedQty, 10);
    saveStore(store);

    const metrics = calculateMetrics(store);
    res.json({ success: true, stockItems: metrics.stockItems });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// PUT Update Opening Stock directly for an item/size
app.put('/api/stock-take/opening', (req, res) => {
  try {
    const { key, openingQty } = req.body;
    const store = getStore();

    if (!store.inventory[key]) {
      return res.status(404).json({ error: 'Stock item not found' });
    }

    store.inventory[key].openingQty = Math.max(0, parseInt(openingQty, 10) || 0);
    saveStore(store);

    const metrics = calculateMetrics(store);
    res.json({ 
      success: true, 
      message: `Updated opening stock for ${store.inventory[key].productName} (${store.inventory[key].size}) to ${store.inventory[key].openingQty}`,
      summary: metrics.summary, 
      stockItems: metrics.stockItems 
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// POST Zero All Stock (Fresh slate for handover)
app.post('/api/stock-take/zero-all', (req, res) => {
  try {
    const store = getStore();
    
    // Set all inventory opening and arrived to 0
    Object.keys(store.inventory).forEach(k => {
      store.inventory[k].openingQty = 0;
      store.inventory[k].arrivedQty = 0;
      store.inventory[k].countedQty = null;
    });

    // Also clear arrivals
    store.arrivals = [];
    
    saveStore(store);

    const metrics = calculateMetrics(store);
    res.json({ 
      success: true, 
      message: 'All inventory reset to 0. You can now enter your true opening stock.',
      summary: metrics.summary, 
      stockItems: metrics.stockItems 
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// POST Add a new clothing/product item
app.post('/api/products', (req, res) => {
  try {
    const { name, category, colour, sizes, costPrice, costPriceVat, sellingPrice, openingStock } = req.body;
    if (!name || !name.trim()) {
      return res.status(400).json({ error: 'Product name is required' });
    }

    let parsedSizes = [];
    if (Array.isArray(sizes)) {
      parsedSizes = sizes.map(s => String(s).trim()).filter(Boolean);
    } else if (typeof sizes === 'string') {
      parsedSizes = sizes.split(',').map(s => s.trim()).filter(Boolean);
    }

    if (parsedSizes.length === 0) {
      parsedSizes = ['One size'];
    }

    const store = getStore();
    
    // Generate clean id
    const baseSlug = name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
    let newId = baseSlug || 'item';
    let counter = 1;
    while (store.products.some(p => p.id === newId)) {
      newId = `${baseSlug}-${counter++}`;
    }

    const cp = Number(costPrice) || 0;
    const cpVat = costPriceVat !== undefined && costPriceVat !== '' ? Number(costPriceVat) : Math.round(cp * 1.15 * 100) / 100;
    const sp = Number(sellingPrice) || 0;

    const newProduct = {
      id: newId,
      name: name.trim(),
      category: category ? category.trim() : 'Clothing',
      colour: colour ? colour.trim() : 'Standard',
      costPrice: cp,
      costPriceVat: cpVat,
      sellingPrice: sp,
      sizes: parsedSizes,
      defaultEst: {}
    };

    store.products.push(newProduct);

    // Initialize inventory for each size
    parsedSizes.forEach(size => {
      const key = `${newId}__${size}`;
      const startingQty = (openingStock && openingStock[size]) ? parseInt(openingStock[size], 10) || 0 : 0;
      store.inventory[key] = {
        productId: newId,
        productName: newProduct.name,
        size: size,
        openingQty: startingQty,
        arrivedQty: 0,
        countedQty: null
      };
    });

    saveStore(store);

    const metrics = calculateMetrics(store);
    res.json({
      success: true,
      message: `Product "${newProduct.name}" added successfully with sizes: ${parsedSizes.join(', ')}`,
      product: newProduct,
      summary: metrics.summary,
      stockItems: metrics.stockItems
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// DELETE a product
app.delete('/api/products/:id', (req, res) => {
  try {
    const store = getStore();
    const idx = store.products.findIndex(p => p.id === req.params.id);
    if (idx === -1) {
      return res.status(404).json({ error: 'Product not found' });
    }

    const deleted = store.products.splice(idx, 1)[0];

    // Remove corresponding inventory keys
    Object.keys(store.inventory).forEach(key => {
      if (store.inventory[key].productId === req.params.id) {
        delete store.inventory[key];
      }
    });

    saveStore(store);
    const metrics = calculateMetrics(store);
    res.json({
      success: true,
      message: `Product "${deleted.name}" removed`,
      summary: metrics.summary,
      stockItems: metrics.stockItems
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// PUT Update Product pricing
app.put('/api/products/:id', (req, res) => {
  try {
    const { costPrice, costPriceVat, sellingPrice } = req.body;
    const store = getStore();
    const product = store.products.find(p => p.id === req.params.id);
    if (!product) {
      return res.status(404).json({ error: 'Product not found' });
    }

    if (costPrice !== undefined) product.costPrice = Number(costPrice);
    if (costPriceVat !== undefined) product.costPriceVat = Number(costPriceVat);
    if (sellingPrice !== undefined) product.sellingPrice = Number(sellingPrice);

    saveStore(store);
    const metrics = calculateMetrics(store);
    res.json({ success: true, product, summary: metrics.summary, stockItems: metrics.stockItems });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// POST Reset to initial template
app.post('/api/reset-data', (req, res) => {
  try {
    if (fs.existsSync(STORE_FILE)) {
      fs.unlinkSync(STORE_FILE);
    }
    const store = getStore();
    const metrics = calculateMetrics(store);
    res.json({ success: true, message: 'Data reset to defaults', summary: metrics.summary, stockItems: metrics.stockItems });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET Export to formatted Excel workbook
app.get('/api/export-excel', async (req, res) => {
  try {
    const store = getStore();
    const metrics = calculateMetrics(store);

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Merch Inventory System';
    workbook.created = new Date();

    // 1. Dashboard Sheet
    const summarySheet = workbook.addWorksheet('Summary & Profit');
    summarySheet.columns = [
      { header: 'Metric', key: 'metric', width: 32 },
      { header: 'Value', key: 'value', width: 22 }
    ];
    summarySheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    summarySheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };

    const s = metrics.summary;
    summarySheet.addRow({ metric: 'Total Revenue (R)', value: s.totalRevenue });
    summarySheet.addRow({ metric: 'Total Cost of Goods Sold (R)', value: s.totalCostOfGoodsSold });
    summarySheet.addRow({ metric: 'Total Net Profit (R)', value: s.totalProfit });
    summarySheet.addRow({ metric: 'Profit Margin (%)', value: s.marginPct.toFixed(2) + '%' });
    summarySheet.addRow({ metric: 'Total Units Sold', value: s.totalUnitsSold });
    summarySheet.addRow({ metric: 'Total Remaining Stock Units', value: s.totalStockRemaining });
    summarySheet.addRow({ metric: 'Stock Valuation at Cost (R)', value: s.totalStockValuationCost });
    summarySheet.addRow({ metric: 'Stock Valuation at Retail (R)', value: s.totalStockValuationRetail });
    summarySheet.addRow({ metric: 'Potential Remaining Profit (R)', value: s.potentialRemainingProfit });

    for (let r = 2; r <= 10; r++) {
      summarySheet.getRow(r).font = { size: 11 };
    }

    // 2. Stock on Hand Sheet
    const stockSheet = workbook.addWorksheet('Stock on Hand');
    stockSheet.columns = [
      { header: 'Product Item', key: 'productName', width: 30 },
      { header: 'Size', key: 'size', width: 12 },
      { header: 'Opening Qty', key: 'openingQty', width: 14 },
      { header: 'Arrived Qty', key: 'arrivedQty', width: 14 },
      { header: 'Total Received', key: 'totalReceived', width: 16 },
      { header: 'Sold to Date', key: 'soldToDate', width: 14 },
      { header: 'Expected on Hand', key: 'stockOnHand', width: 18 },
      { header: 'Counted (Optional)', key: 'countedQty', width: 18 },
      { header: 'Variance', key: 'variance', width: 14 },
      { header: 'Cost Price (R)', key: 'costPrice', width: 15 },
      { header: 'Selling Price (R)', key: 'sellingPrice', width: 16 },
      { header: 'Unit Profit (R)', key: 'profitPerUnit', width: 15 },
      { header: 'Status', key: 'status', width: 15 }
    ];
    stockSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    stockSheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };

    metrics.stockItems.forEach(item => {
      stockSheet.addRow({
        productName: item.productName,
        size: item.size,
        openingQty: item.openingQty,
        arrivedQty: item.arrivedQty,
        totalReceived: item.totalReceived,
        soldToDate: item.soldToDate,
        stockOnHand: item.stockOnHand,
        countedQty: item.countedQty !== null ? item.countedQty : '',
        variance: item.variance !== null ? item.variance : '',
        costPrice: item.costPrice,
        sellingPrice: item.sellingPrice,
        profitPerUnit: item.profitPerUnit,
        status: item.status
      });
    });

    // 3. Sales Log Sheet
    const salesSheet = workbook.addWorksheet('Sales Log');
    salesSheet.columns = [
      { header: 'Date', key: 'date', width: 14 },
      { header: 'Buyer', key: 'buyer', width: 24 },
      { header: 'Item', key: 'productName', width: 30 },
      { header: 'Size', key: 'size', width: 10 },
      { header: 'Qty', key: 'qty', width: 10 },
      { header: 'Unit Price (R)', key: 'unitPrice', width: 14 },
      { header: 'Total Amount (R)', key: 'amount', width: 16 },
      { header: 'Cost Price (R)', key: 'costPrice', width: 14 },
      { header: 'Profit (R)', key: 'profit', width: 14 },
      { header: 'Notes', key: 'notes', width: 25 }
    ];
    salesSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    salesSheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D4ED8' } };

    store.sales.forEach(sale => {
      salesSheet.addRow({
        date: sale.date,
        buyer: sale.buyer,
        productName: sale.productName,
        size: sale.size,
        qty: sale.qty,
        unitPrice: sale.unitPrice,
        amount: sale.amount,
        costPrice: sale.costPrice,
        profit: sale.profit,
        notes: sale.notes || ''
      });
    });

    // 4. Arrivals Log Sheet
    const arrivalsSheet = workbook.addWorksheet('Arrivals Log');
    arrivalsSheet.columns = [
      { header: 'Date', key: 'date', width: 14 },
      { header: 'Product Item', key: 'productName', width: 30 },
      { header: 'Total Units Received', key: 'totalUnits', width: 20 },
      { header: 'Sizes Breakdown', key: 'breakdown', width: 35 },
      { header: 'Notes', key: 'notes', width: 25 }
    ];
    arrivalsSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    arrivalsSheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB45309' } };

    store.arrivals.forEach(arr => {
      const breakdownStr = Object.entries(arr.sizeBreakdown || {})
        .filter(([_, q]) => Number(q) > 0)
        .map(([sz, q]) => `${sz}: ${q}`)
        .join(', ');
      arrivalsSheet.addRow({
        date: arr.date,
        productName: arr.productName,
        totalUnits: arr.totalUnits,
        breakdown: breakdownStr,
        notes: arr.notes || ''
      });
    });

    const filename = `Merch_Inventory_${new Date().toISOString().split('T')[0]}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`=================================================`);
  console.log(` Merch Management Portal running at:`);
  console.log(` Local:   http://localhost:${PORT}`);
  console.log(` Network: http://0.0.0.0:${PORT}`);
  console.log(`=================================================`);
});
