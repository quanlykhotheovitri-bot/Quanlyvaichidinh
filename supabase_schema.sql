-- Create tables for Ortholite Inventory Management

-- Products table
CREATE TABLE products (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  sku TEXT NOT NULL,
  name TEXT NOT NULL,
  category TEXT,
  unit TEXT,
  minStock NUMERIC DEFAULT 0,
  lotNo TEXT,
  ghiChu TEXT,
  designationCode TEXT,
  loaiChiDinh TEXT,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Transactions table
CREATE TABLE transactions (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  productId UUID REFERENCES products(id) ON DELETE CASCADE,
  type TEXT CHECK (type IN ('inbound', 'outbound')),
  quantity NUMERIC NOT NULL,
  date TIMESTAMPTZ NOT NULL,
  partner TEXT,
  loaiChiDinh TEXT,
  lotNo TEXT,
  ghiChu TEXT,
  designationCode TEXT,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Customers table
CREATE TABLE customers (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  code TEXT NOT NULL UNIQUE,
  name TEXT NOT NULL,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Delivery Notes table
CREATE TABLE delivery_notes (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
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
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Enable Row Level Security (RLS)
ALTER TABLE products ENABLE ROW LEVEL SECURITY;
ALTER TABLE transactions ENABLE ROW LEVEL SECURITY;
ALTER TABLE customers ENABLE ROW LEVEL SECURITY;
ALTER TABLE delivery_notes ENABLE ROW LEVEL SECURITY;

-- Create policies for public access (adjust as needed for production)
CREATE POLICY "Public Access" ON products FOR ALL USING (true);
CREATE POLICY "Public Access" ON transactions FOR ALL USING (true);
CREATE POLICY "Public Access" ON customers FOR ALL USING (true);
CREATE POLICY "Public Access" ON delivery_notes FOR ALL USING (true);
