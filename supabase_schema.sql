-- Create products table
CREATE TABLE IF NOT EXISTS products (
  id TEXT PRIMARY KEY,
  sku TEXT NOT NULL,
  name TEXT NOT NULL,
  category TEXT,
  unit TEXT,
  minStock NUMERIC DEFAULT 0,
  lotNo TEXT,
  ghiChu TEXT,
  designationCode TEXT,
  loaiChiDinh TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create transactions table
CREATE TABLE IF NOT EXISTS transactions (
  id TEXT PRIMARY KEY,
  productId TEXT REFERENCES products(id) ON DELETE CASCADE,
  type TEXT CHECK (type IN ('inbound', 'outbound')),
  quantity NUMERIC NOT NULL,
  date TEXT NOT NULL,
  partner TEXT,
  loaiChiDinh TEXT,
  lotNo TEXT,
  ghiChu TEXT,
  designationCode TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create customers table
CREATE TABLE IF NOT EXISTS customers (
  id TEXT PRIMARY KEY,
  code TEXT NOT NULL,
  name TEXT NOT NULL,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create delivery_notes table
CREATE TABLE IF NOT EXISTS delivery_notes (
  id TEXT PRIMARY KEY,
  no INTEGER,
  ovnSaleOrder TEXT,
  ovnProductionOrder TEXT,
  item TEXT,
  materialName TEXT,
  unit TEXT,
  qtyErp NUMERIC,
  actualQty NUMERIC,
  lotNo TEXT,
  actualIssuedQty NUMERIC,
  remark TEXT,
  brand TEXT,
  customerCode TEXT,
  finalDestination TEXT,
  noCode TEXT,
  location TEXT,
  stock TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create location_entries table
CREATE TABLE IF NOT EXISTS location_entries (
  id TEXT PRIMARY KEY,
  qrcode TEXT NOT NULL,
  sku TEXT NOT NULL,
  partner TEXT,
  date TEXT,
  location TEXT,
  note TEXT,
  quantity NUMERIC DEFAULT 1,
  type TEXT CHECK (type IN ('input', 'inventory')),
  scanType TEXT CHECK (scanType IN ('INPUT', 'OUTPUT')),
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create saved_delivery_notes table
CREATE TABLE IF NOT EXISTS saved_delivery_notes (
  id TEXT PRIMARY KEY,
  date TEXT NOT NULL,
  items JSONB NOT NULL,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create delivery_note_header table
CREATE TABLE IF NOT EXISTS delivery_note_header (
  id TEXT PRIMARY KEY DEFAULT 'current',
  docCode TEXT,
  dept TEXT,
  toName TEXT,
  date TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Enable Row Level Security (RLS) for all tables
ALTER TABLE products ENABLE ROW LEVEL SECURITY;
ALTER TABLE transactions ENABLE ROW LEVEL SECURITY;
ALTER TABLE customers ENABLE ROW LEVEL SECURITY;
ALTER TABLE delivery_notes ENABLE ROW LEVEL SECURITY;
ALTER TABLE location_entries ENABLE ROW LEVEL SECURITY;
ALTER TABLE saved_delivery_notes ENABLE ROW LEVEL SECURITY;
ALTER TABLE delivery_note_header ENABLE ROW LEVEL SECURITY;

-- Create policies to allow all access (for development)
-- In a production app, you should restrict these policies
CREATE POLICY "Allow all access to products" ON products FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to transactions" ON transactions FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to customers" ON customers FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to delivery_notes" ON delivery_notes FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to location_entries" ON location_entries FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to saved_delivery_notes" ON saved_delivery_notes FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to delivery_note_header" ON delivery_note_header FOR ALL USING (true) WITH CHECK (true);
