import { format } from 'date-fns';

export const INITIAL_PRODUCTS: any[] = [
  { 
    id: '1', 
    sku: 'AWB-022961', 
    name: '19-4014TCX YH-S817A EPM5 (100% REC.PES) 44" 220G', 
    category: 'Vải', 
    unit: 'Yds', 
    minStock: 10,
    lotNo: 'cut_06/03/25Lot',
    ghiChu: '',
    designationCode: 'C0211',
    loaiChiDinh: ''
  },
];

export const INITIAL_TRANSACTIONS: any[] = [
  { 
    id: 't1', 
    productId: '1', 
    type: 'inbound', 
    quantity: 22, 
    date: '06/03/2025', 
    partner: 'Supplier A', 
    loaiChiDinh: '',
    lotNo: 'cut_06/03/25Lot',
    ghiChu: '',
    designationCode: 'C0211'
  },
  { 
    id: 't2', 
    productId: '1', 
    type: 'outbound', 
    quantity: 22, 
    date: '06/03/2025', 
    partner: 'Customer B', 
    loaiChiDinh: '',
    lotNo: 'cut_06/03/25Lot',
    ghiChu: '',
    designationCode: 'C0211'
  },
];

export const INITIAL_CUSTOMERS: any[] = [
  { id: 'c1', code: 'C0299', name: 'ZHANCHENG INTERNATIONAL TRADE COMPANY LIMITED' },
  { id: 'c2', code: 'C0319', name: 'A PLUS A FOOTWEAR TRADING LIMITED' },
  { id: 'c3', code: 'C0161', name: 'A2DS SCHUHFABRIK PVT LTD' },
  { id: 'c4', code: 'C0290', name: 'ADIANA VIETNAM FOOTWEAR COMPANY LIMITED/NICE ELITE INTERNATIONAL LIMITED' },
  { id: 'c5', code: 'C0286', name: 'ALENA VIETNAM FOOTWEAR COMPANY LIMITED / UNITED WELL TRADING LIMITED' },
];
