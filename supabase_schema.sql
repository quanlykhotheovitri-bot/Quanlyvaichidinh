-- Dọn dẹp các bảng cũ bị lỗi (CẢNH BÁO: Xoá dữ liệu cũ để set-up lại từ đầu, nếu có dữ liệu cực kỳ quan trọng không được xoá thì nhắn lại mình nhé!)
DROP TABLE IF EXISTS delivery_note_header;
DROP TABLE IF EXISTS saved_delivery_notes;
DROP TABLE IF EXISTS location_entries;
DROP TABLE IF EXISTS delivery_notes;
DROP TABLE IF EXISTS customers;
DROP TABLE IF EXISTS transactions;
DROP TABLE IF EXISTS products;

-- 1. Create products table
CREATE TABLE products (
  id TEXT PRIMARY KEY,
  sku TEXT NOT NULL,
  name TEXT NOT NULL,
  category TEXT,
  unit TEXT,
  minstock NUMERIC DEFAULT 0,
  lotno TEXT,
  ghichu TEXT,
  designationcode TEXT,
  loaichidinh TEXT,
  rpro TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 2. Create transactions table
CREATE TABLE transactions (
  id TEXT PRIMARY KEY,
  productid TEXT REFERENCES products(id) ON DELETE CASCADE,
  type TEXT CHECK (type IN ('inbound', 'outbound')),
  quantity NUMERIC NOT NULL,
  date TEXT NOT NULL,
  partner TEXT,
  loaichidinh TEXT,
  lotno TEXT,
  ghichu TEXT,
  designationcode TEXT,
  rpro TEXT,
  updatedate TEXT,
  isdeleted BOOLEAN DEFAULT false,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 3. Create customers table
CREATE TABLE customers (
  id TEXT PRIMARY KEY,
  code TEXT NOT NULL,
  name TEXT NOT NULL,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 4. Create delivery_notes table
CREATE TABLE delivery_notes (
  id TEXT PRIMARY KEY,
  no INTEGER,
  ovnsaleorder TEXT,
  ovnproductionorder TEXT,
  item TEXT,
  materialname TEXT,
  unit TEXT,
  qtyerp NUMERIC,
  actualqty NUMERIC,
  lotno TEXT,
  actualissuedqty NUMERIC,
  remark TEXT,
  brand TEXT,
  customercode TEXT,
  finaldestination TEXT,
  nocode TEXT,
  location TEXT,
  stock TEXT,
  loaichidinh TEXT,
  assignedlots JSONB,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 5. Create location_entries table
CREATE TABLE location_entries (
  id TEXT PRIMARY KEY,
  qrcode TEXT NOT NULL,
  sku TEXT NOT NULL,
  partner TEXT,
  date TEXT,
  location TEXT,
  note TEXT,
  quantity NUMERIC DEFAULT 1,
  type TEXT CHECK (type IN ('input', 'inventory')),
  scantype TEXT CHECK (scantype IN ('INPUT', 'OUTPUT')),
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 6. Create saved_delivery_notes table
CREATE TABLE saved_delivery_notes (
  id TEXT PRIMARY KEY,
  date TEXT NOT NULL,
  items JSONB NOT NULL,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 7. Create delivery_note_header table
CREATE TABLE delivery_note_header (
  id TEXT PRIMARY KEY DEFAULT 'current',
  doccode TEXT,
  dept TEXT,
  toname TEXT,
  date TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Kích hoạt quyền truy cập
ALTER TABLE products ENABLE ROW LEVEL SECURITY;
ALTER TABLE transactions ENABLE ROW LEVEL SECURITY;
ALTER TABLE customers ENABLE ROW LEVEL SECURITY;
ALTER TABLE delivery_notes ENABLE ROW LEVEL SECURITY;
ALTER TABLE location_entries ENABLE ROW LEVEL SECURITY;
ALTER TABLE saved_delivery_notes ENABLE ROW LEVEL SECURITY;
ALTER TABLE delivery_note_header ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Allow all access to products" ON products FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to transactions" ON transactions FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to customers" ON customers FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to delivery_notes" ON delivery_notes FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to location_entries" ON location_entries FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to saved_delivery_notes" ON saved_delivery_notes FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "Allow all access to delivery_note_header" ON delivery_note_header FOR ALL USING (true) WITH CHECK (true);
