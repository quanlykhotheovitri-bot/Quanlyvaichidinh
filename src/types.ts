export interface Product {
  id: string;
  sku: string; // Mã Hàng
  name: string; // Tên hàng
  category: string;
  unit: string;
  minStock: number;
  lotNo?: string;
  ghiChu?: string;
  designationCode?: string; // Mã chỉ định
  loaiChiDinh?: string; // Loại chỉ định
  rpro?: string; // OVN Production Order
  created_at?: string;
}

export interface Transaction {
  id: string;
  productId: string;
  type: 'inbound' | 'outbound';
  quantity: number;
  date: string;
  partner: string;
  loaiChiDinh?: string;
  lotNo?: string;
  ghiChu?: string;
  designationCode?: string;
  rpro?: string;
  created_at?: string;
}

export interface Customer {
  id: string;
  code: string;
  name: string;
  created_at?: string;
}

export interface DeliveryNoteItem {
  id: string;
  no: number;
  ovnSaleOrder: string;
  ovnProductionOrder: string;
  item: string;
  materialName: string;
  unit: string;
  qtyErp: number;
  actualQty?: number;
  lotNo: string;
  actualIssuedQty?: number;
  remark: string;
  brand: string;
  customerCode: string;
  finalDestination: string;
  noCode: string;
  location: string;
  stock: string;
  loaiChiDinh?: string;
  created_at?: string;
}

export interface InventoryItem extends Product {
  productId: string;
  totalInbound: number;
  totalOutbound: number;
  currentStock: number;
  tempStock?: number;
  inboundDate?: number;
}

export interface LocationEntry {
  id: string;
  qrcode: string;
  sku: string;
  partner: string;
  date: string;
  location: string;
  note: string;
  quantity: number;
  type: 'input' | 'inventory';
  scanType?: 'INPUT' | 'OUTPUT';
  created_at?: string;
}
