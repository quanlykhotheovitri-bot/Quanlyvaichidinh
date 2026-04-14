import React, { useState, useMemo, useRef, useCallback, useEffect } from 'react';
import { 
  LayoutDashboard, 
  ArrowDownToLine, 
  ArrowUpFromLine, 
  Package, 
  Plus, 
  FileUp, 
  Search, 
  AlertTriangle,
  ChevronRight,
  Download,
  Filter,
  Users,
  Trash2,
  Edit2,
  CheckSquare,
  Square,
  FileText,
  Printer,
  MapPin,
  Save,
  WifiOff,
  AlertCircle,
  CheckCircle2,
  X,
  Menu
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import Papa from 'papaparse';
import { format, parse, isValid } from 'date-fns';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  Cell
} from 'recharts';

import { cn } from './lib/utils';
import { Product, Transaction, InventoryItem, Customer, DeliveryNoteItem, LocationEntry } from './types';
import { INITIAL_PRODUCTS, INITIAL_TRANSACTIONS, INITIAL_CUSTOMERS } from './constants';
import { api } from './lib/api';
import { isSupabaseConfigured as INITIAL_SUPABASE_CONFIGURED } from './lib/supabase';

import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

type Tab = 'dashboard' | 'inbound' | 'outbound' | 'inventory' | 'customers' | 'deliveryNote' | 'location';

export default function App() {
  const [activeTab, setActiveTab] = useState<Tab>('dashboard');
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [products, setProducts] = useState<Product[]>([]);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [locationEntries, setLocationEntries] = useState<LocationEntry[]>([]);
  const [locationInventoryEntries, setLocationInventoryEntries] = useState<LocationEntry[]>([]);
  const [deliveryNotes, setDeliveryNotes] = useState<DeliveryNoteItem[]>([]);
  const [savedDeliveryNotes, setSavedDeliveryNotes] = useState<{id: string, date: string, items: DeliveryNoteItem[]}[]>([]);
  const [deliveryNoteHeader, setDeliveryNoteHeader] = useState({
    documentCode: 'WH.F-004/P-01',
    dept: 'SX 5',
    to: '',
    date: ''
  });
  const [deliveryNoteSubTab, setDeliveryNoteSubTab] = useState<'draft' | 'preview' | 'history'>('preview');
  const [locationSubTab, setLocationSubTab] = useState<'input' | 'output' | 'inventory'>('input');
  const [locationSearch, setLocationSearch] = useState('');
  const [currentLocation, setCurrentLocation] = useState('');
  const [scanMode, setScanMode] = useState<'INPUT' | 'OUTPUT'>('INPUT');
  const [searchQuery, setSearchQuery] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);
  
  const [selectedRows, setSelectedRows] = useState<string[]>([]);
  const [editingId, setEditingId] = useState<string | null>(null);

  const [isProductModalOpen, setIsProductModalOpen] = useState(false);
  const [isTransactionModalOpen, setIsTransactionModalOpen] = useState(false);
  const [isCustomerModalOpen, setIsCustomerModalOpen] = useState(false);
  const [isDeleteConfirmOpen, setIsDeleteConfirmOpen] = useState(false);
  const [isDeliveryNoteDeleteConfirmOpen, setIsDeliveryNoteDeleteConfirmOpen] = useState(false);
  const [deliveryNoteDeleteId, setDeliveryNoteDeleteId] = useState<string | null>(null);
  const [isDeliveryNoteEditModalOpen, setIsDeliveryNoteEditModalOpen] = useState(false);
  const [editingDeliveryNoteId, setEditingDeliveryNoteId] = useState<string | null>(null);
  const [tempDeliveryNoteItem, setTempDeliveryNoteItem] = useState<Partial<DeliveryNoteItem>>({});
  const [isLocationModalOpen, setIsLocationModalOpen] = useState(false);
  const [newLocationEntry, setNewLocationEntry] = useState<Partial<LocationEntry>>({
    qrcode: '', sku: '', partner: '', date: '', location: '', note: '', quantity: 1
  });
  const [deleteTarget, setDeleteTarget] = useState<{ id: string | 'bulk' | 'wipe_all_data', type: 'product' | 'transaction' | 'customer' | 'location' | 'savedDeliveryNote' | 'all', qrcode?: string } | null>(null);
  const [notification, setNotification] = useState<{ message: string, type: 'success' | 'error' } | null>(null);

  const showNotification = (message: string, type: 'success' | 'error' = 'success') => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 5000);
  };
  const [isSupabaseConfigured, setIsSupabaseConfigured] = useState(INITIAL_SUPABASE_CONFIGURED);
  const [isSetupModalOpen, setIsSetupModalOpen] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [hasLoadedData, setHasLoadedData] = useState(false);
  const [isOnline, setIsOnline] = useState(navigator.onLine);
  const syncQueue = useRef<Promise<any>>(Promise.resolve());

  const queueSync = useCallback((syncFn: () => Promise<any>) => {
    syncQueue.current = syncQueue.current.then(syncFn).catch(err => {
      console.error('Sync error in queue:', err);
    });
  }, []);

  const [newProduct, setNewProduct] = useState<Partial<Product> & { quantity?: number }>({
    sku: '', name: '', category: '', unit: '', minStock: 0, lotNo: '', ghiChu: '', designationCode: '', loaiChiDinh: '', quantity: 0
  });
  const [newTransaction, setNewTransaction] = useState<Partial<Transaction>>({
    productId: '', type: 'inbound', quantity: 0, date: format(new Date(), 'dd/MM/yyyy'), partner: '', loaiChiDinh: '', lotNo: '', ghiChu: '', designationCode: ''
  });
  const [newCustomer, setNewCustomer] = useState<Partial<Customer>>({
    code: '', name: ''
  });
  const [scanInput, setScanInput] = useState('');

  useEffect(() => {
    setSelectedRows([]);
  }, [activeTab, deliveryNoteSubTab]);

  const generateId = useCallback(() => {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) {
      return crypto.randomUUID();
    }
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
      var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }, []);

  const extractDateFromLot = useCallback((lot: string): number => {
    const slashMatch = lot.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (slashMatch) {
      let [_, d, m, y] = slashMatch;
      const year = y.length === 2 ? 2000 + parseInt(y) : parseInt(y);
      return new Date(year, parseInt(m) - 1, parseInt(d)).getTime();
    }
    
    const hyphenMatch = lot.match(/(\d{1,2})-(\d{1,2})-(\d{2,4})/);
    if (hyphenMatch) {
      let [_, d, m, y] = hyphenMatch;
      const year = y.length === 2 ? 2000 + parseInt(y) : parseInt(y);
      return new Date(year, parseInt(m) - 1, parseInt(d)).getTime();
    }

    const isoMatch = lot.match(/(\d{2,4})-(\d{1,2})-(\d{1,2})/);
    if (isoMatch) {
      let [_, y, m, d] = isoMatch;
      const year = y.length === 2 ? 2000 + parseInt(y) : parseInt(y);
      return new Date(year, parseInt(m) - 1, parseInt(d)).getTime();
    }

    return 9999999999999;
  }, []);

  const normalizeDateForMatching = useCallback((lot: string): string => {
    if (!lot) return '';
    const cleanLot = lot.trim();
    
    const slashMatch = cleanLot.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (slashMatch) {
      let [_, d, m, y] = slashMatch;
      const year = y.length === 2 ? '20' + y : y;
      const day = d.padStart(2, '0');
      const month = m.padStart(2, '0');
      return `${day}/${month}/${year}`;
    }
    
    const hyphenMatch = cleanLot.match(/(\d{1,2})-(\d{1,2})-(\d{2,4})/);
    if (hyphenMatch) {
      let [_, d, m, y] = hyphenMatch;
      const year = y.length === 2 ? '20' + y : y;
      const day = d.padStart(2, '0');
      const month = m.padStart(2, '0');
      return `${day}/${month}/${year}`;
    }

    const isoMatch = cleanLot.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (isoMatch) {
      let [_, y, m, d] = isoMatch;
      const day = d.padStart(2, '0');
      const month = m.padStart(2, '0');
      return `${day}/${month}/${y}`;
    }

    // Handle YYYYMMDD
    if (/^\d{8}$/.test(cleanLot)) {
      const y = cleanLot.substring(0, 4);
      const m = cleanLot.substring(4, 6);
      const d = cleanLot.substring(6, 8);
      // Check if it looks like a valid year (e.g. 20xx)
      if (parseInt(y) > 1900 && parseInt(y) < 2100 && parseInt(m) <= 12 && parseInt(d) <= 31) {
        return `${d}/${m}/${y}`;
      }
      // Or maybe DDMMYYYY
      const d2 = cleanLot.substring(0, 2);
      const m2 = cleanLot.substring(2, 4);
      const y2 = cleanLot.substring(4, 8);
      if (parseInt(y2) > 1900 && parseInt(y2) < 2100 && parseInt(m2) <= 12 && parseInt(d2) <= 31) {
        return `${d2}/${m2}/${y2}`;
      }
    }

    return cleanLot; // Return original if no match, instead of empty string
  }, []);

  const resolveByPriority1 = useCallback((sku: string, rpro: string, inventoryRows: InventoryItem[]) => {
    const targetRpro = rpro.replace(/\s+/g, '').toLowerCase();
    if (!targetRpro) return [];
    
    return inventoryRows.filter(item => {
      const itemSku = item.sku.toLowerCase().trim();
      if (itemSku !== sku) return false;
      
      const itemRpro = (item.rpro || '').replace(/\s+/g, '').toLowerCase();
      if (itemRpro === targetRpro) return true;

      // Tách mã chỉ định bằng nhiều loại dấu phân cách: /, dấu phẩy, hoặc khoảng trắng
      const designationCode = (item.designationCode || '').toLowerCase().trim();
      const codes = designationCode.split(/[\/,\s]+/).map(c => c.trim()).filter(Boolean);
      
      return codes.includes(targetRpro);
    });
  }, []);

  const resolveByPriority2 = useCallback((sku: string, no: string, inventoryRows: InventoryItem[]) => {
    const targetNo = no.replace(/\s+/g, '').toLowerCase();
    // Chỉ xét nếu mã khách hàng bắt đầu bằng chữ C
    if (!targetNo || !targetNo.startsWith('c')) return [];
    
    return inventoryRows.filter(item => {
      const itemSku = item.sku.toLowerCase().trim();
      if (itemSku !== sku) return false;
      
      const designationCode = (item.designationCode || '').toLowerCase().trim();
      const codes = designationCode.split(/[\/,\s]+/).map(c => c.trim()).filter(Boolean);
      
      return codes.includes(targetNo);
    });
  }, []);

  const resolveByPriority3 = useCallback((sku: string, inventoryRows: InventoryItem[]) => {
    return inventoryRows.filter(item => {
      const itemSku = item.sku.toLowerCase().trim();
      if (itemSku !== sku) return false;
      
      const designationCode = (item.designationCode || '').trim();
      return designationCode === '';
    });
  }, []);

  const findAssignedFabricLot = useCallback((issueRow: { sku: string, rpro: string, no: string }, inventoryRows: InventoryItem[]) => {
    const sku = issueRow.sku.toLowerCase().trim();
    
    // Priority 1
    let matches = resolveByPriority1(sku, issueRow.rpro, inventoryRows);
    
    // Priority 2
    if (matches.length === 0) {
      matches = resolveByPriority2(sku, issueRow.no, inventoryRows);
    }
    
    // Priority 3
    if (matches.length === 0) {
      matches = resolveByPriority3(sku, inventoryRows);
    }
    
    // Filter by stock > 0
    matches = matches.filter(m => (m.tempStock !== undefined ? m.tempStock : m.currentStock) > 0);
    
    // Sort FIFO: 1) inboundDate, 2) ID
    matches.sort((a, b) => {
      const dateA = a.inboundDate || 9999999999999;
      const dateB = b.inboundDate || 9999999999999;
      if (dateA !== dateB) return dateA - dateB;
      return a.id.localeCompare(b.id);
    });
    
    return matches;
  }, [resolveByPriority1, resolveByPriority2, resolveByPriority3]);

  React.useEffect(() => {
    setSelectedRows([]);
  }, [activeTab, locationSubTab]);

  const loadData = async () => {
    try {
      const [dbProducts, dbTransactions, dbCustomers, dbDeliveryNotes, dbLocationEntries, dbSavedDeliveryNotes, dbHeader] = await Promise.all([
        api.products.getAll(),
        api.transactions.getAll(),
        api.customers.getAll(),
        api.deliveryNotes.getAll(),
        api.locationEntries.getAll(),
        api.savedDeliveryNotes.getAll(),
        api.deliveryNoteHeader.get()
      ]);

      if (dbProducts.length > 0) {
        setProducts(dbProducts);
      }

      if (dbTransactions.length > 0) {
        setTransactions(dbTransactions);
      }

      if (dbCustomers.length > 0) {
        setCustomers(dbCustomers);
      }

      if (dbDeliveryNotes.length > 0) setDeliveryNotes(dbDeliveryNotes);
      if (dbLocationEntries.length > 0) {
        setLocationEntries(dbLocationEntries.filter(e => e.type === 'input' || !e.type));
        const inventoryEntries = dbLocationEntries.filter(e => e.type === 'inventory');
        const grouped = new Map<string, LocationEntry>();
        inventoryEntries.forEach(entry => {
          const key = `${entry.qrcode}|${entry.location || ''}`;
          if (!grouped.has(key)) {
            grouped.set(key, { ...entry });
          } else {
            const existing = grouped.get(key)!;
            if (!existing.sku) existing.sku = entry.sku;
            if (!existing.partner) existing.partner = entry.partner;
            if (!existing.date) existing.date = entry.date;
            if (!existing.note) existing.note = entry.note;
          }
        });
        setLocationInventoryEntries(Array.from(grouped.values()));
      }
      if (dbSavedDeliveryNotes.length > 0) setSavedDeliveryNotes(dbSavedDeliveryNotes);
      if (dbHeader) {
        setDeliveryNoteHeader({
          documentCode: dbHeader.docCode || 'WH.F-004/P-01',
          dept: dbHeader.dept || 'SX 5',
          to: dbHeader.to || '',
          date: dbHeader.date || format(new Date(), 'dd/MM/yyyy')
        });
      }
      setHasLoadedData(true);
    } catch (error) {
      console.error('Error loading data from Supabase:', error);
    }
  };

  React.useEffect(() => {
    loadData();
  }, []);

  useEffect(() => {
    if (!hasLoadedData || !isSupabaseConfigured) return;

    const timeoutId = setTimeout(() => {
      if (!isOnline) {
        console.log('App is offline, skipping sync until network is back. Data is safely cached locally.');
        return;
      }

      queueSync(async () => {
        setIsSaving(true);
        try {
          await Promise.all([
            api.products.upsertAll(products),
            api.transactions.upsertAll(transactions),
            api.customers.upsertAll(customers),
            api.locationEntries.upsertAll([...locationEntries, ...locationInventoryEntries]),
            api.deliveryNotes.upsertAll(deliveryNotes),
            api.savedDeliveryNotes.upsertAll(savedDeliveryNotes),
            api.deliveryNoteHeader.upsert({
              docCode: deliveryNoteHeader.documentCode,
              dept: deliveryNoteHeader.dept,
              to: deliveryNoteHeader.to,
              date: deliveryNoteHeader.date
            })
          ]);
        } catch (error) {
          console.error('Lỗi khi tự động lưu dữ liệu:', error);
        } finally {
          setIsSaving(false);
        }
      });
    }, 3000);

    return () => clearTimeout(timeoutId);
  }, [hasLoadedData, isSupabaseConfigured, isOnline, products, transactions, customers, locationEntries, locationInventoryEntries, deliveryNotes, savedDeliveryNotes, deliveryNoteHeader]);

  useEffect(() => {
    const handleOnline = async () => {
      setIsOnline(true);
      if (hasLoadedData && isSupabaseConfigured) {
        setIsSaving(true);
        try {
          await Promise.all([
            api.products.upsertAll(products),
            api.transactions.upsertAll(transactions),
            api.customers.upsertAll(customers),
            api.locationEntries.upsertAll([...locationEntries, ...locationInventoryEntries]),
            api.deliveryNotes.upsertAll(deliveryNotes),
            api.savedDeliveryNotes.upsertAll(savedDeliveryNotes),
            api.deliveryNoteHeader.upsert({
              docCode: deliveryNoteHeader.documentCode,
              dept: deliveryNoteHeader.dept,
              to: deliveryNoteHeader.to,
              date: deliveryNoteHeader.date
            })
          ]);
          showNotification('Đã kết nối lại Internet và đồng bộ dữ liệu thành công!');
        } catch (error) {
          console.error('Lỗi khi đồng bộ lúc có mạng lại:', error);
          showNotification('Lỗi khi đồng bộ dữ liệu.', 'error');
        } finally {
          setIsSaving(false);
        }
      }
    };
    
    const handleOffline = () => {
      setIsOnline(false);
      showNotification('Đã mất kết nối Internet. Dữ liệu đang được lưu cục bộ.', 'error');
    };

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
    };
  }, [hasLoadedData, isSupabaseConfigured, products, transactions, customers, locationEntries, locationInventoryEntries, deliveryNotes, savedDeliveryNotes, deliveryNoteHeader]);

  useEffect(() => {
    if (isSupabaseConfigured) {
      api.deliveryNoteHeader.upsert({
        docCode: deliveryNoteHeader.documentCode,
        dept: deliveryNoteHeader.dept,
        to: deliveryNoteHeader.to,
        date: deliveryNoteHeader.date
      }).catch(err => console.error('Error syncing delivery note header:', err));
    }
  }, [deliveryNoteHeader, isSupabaseConfigured]);

  const handleDeleteDeliveryNoteItem = (id: string) => {
    setDeliveryNotes(prev => {
      const filtered = prev.filter(item => item.id !== id);
      const updatedNotes = recalculateActualQty(filtered);
      api.deliveryNotes.upsertAll(updatedNotes).catch(error => {
        console.error('Error syncing delivery notes:', error);
      });
      return updatedNotes;
    });
    showNotification('Đã xóa dòng khỏi phiếu giao hàng', 'success');
  };

  const confirmDeleteDeliveryNoteItem = async () => {
    if (deliveryNoteDeleteId !== null) {
      setDeliveryNotes(prev => {
        const updatedNotes = prev.filter(item => item.id !== deliveryNoteDeleteId);
        api.deliveryNotes.upsertAll(updatedNotes).catch(error => {
          console.error('Error syncing delivery notes:', error);
          showNotification('Không thể đồng bộ phiếu giao hàng sau khi xóa.', 'error');
        });
        return updatedNotes;
      });
      setDeliveryNoteDeleteId(null);
      setIsDeliveryNoteDeleteConfirmOpen(false);
    }
  };

  const handleEditDeliveryNoteItemClick = (id: string) => {
    const item = deliveryNotes.find(dn => dn.id === id);
    if (item) {
      setEditingDeliveryNoteId(id);
      setTempDeliveryNoteItem(item);
      setIsDeliveryNoteEditModalOpen(true);
    }
  };

  const determineLoaiChiDinh = (itemCode: string, prodOrder: string, saleOrder: string, noCode: string, currentLoai: string) => {
    if (currentLoai && currentLoai !== 'NORMAL') return currentLoai;
    
    const match = inventory.find(inv => 
      inv.sku.toLowerCase().trim() === itemCode.toLowerCase().trim() && 
      (
        (prodOrder && (inv.designationCode || '').toLowerCase().includes(prodOrder.toLowerCase().trim())) ||
        (saleOrder && (inv.designationCode || '').toLowerCase().includes(saleOrder.toLowerCase().trim())) ||
        (noCode && (inv.designationCode || '').toLowerCase().includes(noCode.toLowerCase().trim()))
      )
    );
    return match?.loaiChiDinh || 'NORMAL';
  };

  const recalculateActualQty = (notes: DeliveryNoteItem[]) => {
    const groups = new Map<string, DeliveryNoteItem[]>();
    notes.forEach(item => {
      const loai = determineLoaiChiDinh(item.item, item.ovnProductionOrder, item.ovnSaleOrder, item.noCode, item.loaiChiDinh);
      const groupKey = `${item.item}|${loai}`;
      if (!groups.has(groupKey)) groups.set(groupKey, []);
      groups.get(groupKey)!.push(item);
    });

    groups.forEach(items => {
      const totalQtyErp = items.reduce((sum, i) => sum + i.qtyErp, 0);
      const targetTotal = Math.ceil(totalQtyErp);
      const diff = targetTotal - totalQtyErp;

      // Find row with largest qtyErp in this group
      let maxIdx = 0;
      let maxVal = -1;
      items.forEach((item, idx) => {
        if (item.qtyErp > maxVal) {
          maxVal = item.qtyErp;
          maxIdx = idx;
        }
      });

      items.forEach((item, idx) => {
        const newActualQty = idx === maxIdx 
          ? Number((item.qtyErp + diff).toFixed(4))
          : item.qtyErp;
        
        if (item.actualQty !== newActualQty) {
          item.actualQty = newActualQty;
          item.assignedLots = [];
          item.actualIssuedQty = 0;
          item.stock = '';
        }
      });
    });
    return [...notes];
  };

  const saveDeliveryNoteItemEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (editingDeliveryNoteId !== null) {
      let updatedNotes = deliveryNotes.map(item => {
        if (item.id === editingDeliveryNoteId) {
          const newItem = { ...item, ...tempDeliveryNoteItem } as DeliveryNoteItem;
          // Nếu thay đổi LotNo hoặc Item, xóa assignedLots để tránh sai lệch khi post
          if (tempDeliveryNoteItem.lotNo !== undefined && tempDeliveryNoteItem.lotNo !== item.lotNo) {
            newItem.assignedLots = [];
            newItem.actualIssuedQty = newItem.actualQty;
          }
          return { ...newItem, location: getLocationByItemAndLot(newItem.item, newItem.lotNo) };
        }
        return item;
      });
      updatedNotes = recalculateActualQty(updatedNotes);
      setDeliveryNotes(updatedNotes);
      try {
        await api.deliveryNotes.upsertAll(updatedNotes);
      } catch (error) {
        console.error('Error syncing delivery notes:', error);
      }
      setIsDeliveryNoteEditModalOpen(false);
      setEditingDeliveryNoteId(null);
    }
  };

  const handleEditDeliveryNoteItem = (index: number, field: keyof DeliveryNoteItem, value: any) => {
    const itemAtId = filteredDeliveryNotes[index].id;
    setDeliveryNotes(prev => {
      let updated = prev.map(item => 
        item.id === itemAtId ? { ...item, [field]: value } : item
      );
      if (field === 'qtyErp') {
        updated = recalculateActualQty(updated);
      }
      if (field === 'item' || field === 'lotNo' || field === 'actualQty') {
        updated = updated.map(item => {
          if (item.id === itemAtId) {
            // Xóa assignedLots khi sửa tay LotNo, Item hoặc Số lượng để tránh sai lệch khi post
            return { 
              ...item, 
              assignedLots: [],
              actualIssuedQty: field === 'actualQty' ? Number(value) : item.actualQty,
              location: field === 'actualQty' ? item.location : getLocationByItemAndLot(item.item, item.lotNo) 
            };
          }
          return item;
        });
      }
      return updated;
    });
  };

  const handlePostDeliveryNote = async () => {
    if (deliveryNotes.length === 0) return;

    const today = format(new Date(), 'dd/MM/yyyy');
    const newTransactions: Transaction[] = [];

    deliveryNotes.forEach(item => {
      // Kiểm tra xem assignedLots có khớp với field lotNo hiện tại không
      const assignedLotStr = (item.assignedLots || []).map(l => l.lotNo).join(', ');
      const isSync = assignedLotStr === item.lotNo;

      if (isSync && item.assignedLots && item.assignedLots.length > 0) {
        // Sử dụng chi tiết Lot đã gán
        item.assignedLots.forEach(lot => {
          const product = products.find(p => p.sku === item.item);
          if (product) {
            newTransactions.push({
              id: generateId(),
              productId: product.id,
              type: 'outbound',
              quantity: lot.qty,
              date: today,
              partner: item.customerCode || item.noCode || 'Unknown',
              lotNo: lot.lotNo,
              ghiChu: 'Xuất từ Phiếu giao nhận',
              designationCode: lot.remark,
              loaiChiDinh: lot.loaiChiDinh
            });
          }
        });
      } else if (item.lotNo && (item.actualIssuedQty || item.actualQty) > 0) {
        // Trường hợp chỉnh sửa tay hoặc fallback: Tách lotNo và tìm tồn kho tương ứng
        const lotList = item.lotNo.split(',').map(l => l.trim()).filter(Boolean);
        const product = products.find(p => p.sku === item.item);
        
        if (product && lotList.length > 0) {
          const totalQty = item.actualIssuedQty || item.actualQty || 0;
          const qtyPerLot = totalQty / lotList.length;

          lotList.forEach(lotName => {
            // Tìm item trong inventory khớp với SKU và LotNo để lấy loaiChiDinh/designationCode chính xác
            const invMatch = inventory.find(inv => 
              inv.sku.toLowerCase().trim() === item.item.toLowerCase().trim() && 
              inv.lotNo === lotName
            );

            newTransactions.push({
              id: generateId(),
              productId: product.id,
              type: 'outbound',
              quantity: qtyPerLot,
              date: today,
              partner: item.customerCode || item.noCode || 'Unknown',
              lotNo: lotName,
              ghiChu: 'Xuất từ Phiếu giao nhận',
              designationCode: invMatch?.designationCode || item.remark || '',
              loaiChiDinh: invMatch?.loaiChiDinh || (item.stock.includes('(') ? item.stock.split('(')[1].replace(')', '') : '')
            });
          });
        }
      }
    });

    if (newTransactions.length > 0) {
      const today = format(new Date(), 'dd/MM/yyyy');
      const newSavedNote = {
        id: generateId(),
        date: today,
        items: [...deliveryNotes]
      };

      setTransactions(prev => [...prev, ...newTransactions]);
      setSavedDeliveryNotes(prev => [...prev, newSavedNote]);
      setDeliveryNotes([]);
      
      try {
        await Promise.all([
          ...newTransactions.map(t => api.transactions.upsert(t)),
          api.savedDeliveryNotes.upsert(newSavedNote),
          api.deliveryNotes.deleteAll()
        ]);
        showNotification('Đã post lệnh xuất kho thành công!');
      } catch (error) {
        console.error('Error syncing post delivery note:', error);
        showNotification('Lỗi khi post lệnh xuất kho.', 'error');
      }
    } else {
      showNotification('Không có dữ liệu thực tế để xuất kho!', 'error');
    }
  };

  const createExcelFile = async (
    data: any[], 
    columns: { header: string, key: string, width: number, alignment?: Partial<ExcelJS.Alignment>, isHighlight?: boolean }[], 
    fileName: string, 
    title: string,
    headerColor: string = 'FF001F3F'
  ) => {
    try {
      showNotification('Đang chuẩn bị file Excel...', 'success');
      
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      // 1. Add Title
      const titleRow = worksheet.addRow(['', title]);
      titleRow.height = 30;
      worksheet.mergeCells(1, 2, 1, columns.length + 1);
      titleRow.getCell(2).font = { name: 'Arial', size: 16, bold: true };
      titleRow.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };

      // 2. Add Metadata (Date)
      const dateRow = worksheet.addRow(['', `Ngày lập báo cáo: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`]);
      worksheet.mergeCells(2, 2, 2, columns.length + 1);
      dateRow.getCell(2).font = { name: 'Arial', size: 10, italic: true };
      dateRow.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };

      worksheet.addRow([]); // Empty row

      // 3. Setup Columns
      worksheet.columns = [
        { width: 5 }, // Column A for spacing/Index
        ...columns.map(col => ({ header: col.header, key: col.key, width: col.width }))
      ];

      // 4. Add Header Row
      const headerRow = worksheet.getRow(4);
      headerRow.values = ['', ...columns.map(c => c.header)];
      headerRow.height = 25;
      headerRow.eachCell((cell, colNumber) => {
        if (colNumber > 1) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: headerColor }
          };
          cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 10 };
          cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        }
      });

      // 5. Add Data Rows
      data.forEach((item, index) => {
        const row = worksheet.addRow(['', ...columns.map(col => item[col.key])]);
        row.eachCell((cell, colNumber) => {
          if (colNumber > 1) {
            const colDef = columns[colNumber - 2];
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
            cell.font = { name: 'Arial', size: 10 };
            cell.alignment = colDef.alignment || { vertical: 'middle', horizontal: 'left' };
            
            if (colDef.isHighlight) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFACC15' } // Yellow-400
              };
            }
          }
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `${fileName}_${format(new Date(), 'ddMMyyyy_HHmm')}.xlsx`);
      showNotification('Xuất file Excel thành công!', 'success');
    } catch (error) {
      console.error('Lỗi khi xuất file Excel:', error);
      showNotification('Lỗi khi tạo file Excel. Vui lòng thử lại.', 'error');
    }
  };

  const handleExportInventory = async () => {
    if (filteredInventory.length === 0) {
      showNotification('Không có dữ liệu tồn kho để xuất!', 'error');
      return;
    }
    const columns = [
      { header: 'Mã Hàng', key: 'sku', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
      { header: 'Tên hàng', key: 'name', width: 35 },
      { header: 'Lot no', key: 'lotNo', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
      { header: 'Số lượng nhập', key: 'totalInbound', width: 15, alignment: { horizontal: 'right' as ExcelJS.Alignment['horizontal'] }, isHighlight: true },
      { header: 'Số lượng xuất', key: 'totalOutbound', width: 15, alignment: { horizontal: 'right' as ExcelJS.Alignment['horizontal'] }, isHighlight: true },
      { header: 'Tồn cuối', key: 'currentStock', width: 15, alignment: { horizontal: 'right' as ExcelJS.Alignment['horizontal'] }, isHighlight: true },
      { header: 'Loại chỉ định', key: 'loaiChiDinh', width: 20 },
      { header: 'Ghi chú', key: 'ghiChu', width: 25 },
      { header: 'Mã chỉ định', key: 'designationCode', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
    ];
    await createExcelFile(filteredInventory, columns, 'BaoCaoTonKho', 'BÁO CÁO TỒN KHO');
  };

  const handleExportTransactions = async () => {
    const isInbound = activeTab === 'inbound';
    const title = isInbound ? 'DANH SÁCH NHẬP KHO' : 'DANH SÁCH XUẤT KHO';
    const fileName = isInbound ? 'DanhSachNhapKho' : 'DanhSachXuatKho';
    const headerColor = isInbound ? 'FF1a5f7a' : 'FFC2410C';

    const rawData = filteredTransactions.filter(t => t.type === activeTab);
    if (rawData.length === 0) {
      showNotification(`Không có dữ liệu ${title.toLowerCase()} để xuất!`, 'error');
      return;
    }

    const columns = [
      { header: 'Mã Hàng', key: 'sku', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
      { header: 'Tên hàng', key: 'name', width: 35 },
      { header: 'Lot no', key: 'lotNo', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
      { header: isInbound ? 'Số lượng nhập' : 'Số lượng xuất', key: 'quantity', width: 15, alignment: { horizontal: 'right' as ExcelJS.Alignment['horizontal'] }, isHighlight: true },
      { header: isInbound ? 'Ngày nhập' : 'Ngày xuất', key: 'date', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
      { header: 'Loại chỉ định', key: 'loaiChiDinh', width: 20 },
      { header: 'Ghi chú', key: 'ghiChu', width: 25 },
      { header: 'Mã chỉ định', key: 'designationCode', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
    ];

    const data = rawData.map(t => {
      const product = products.find(p => p.id === t.productId);
      return {
        ...t,
        sku: product?.sku || 'N/A',
        name: product?.name || 'N/A'
      };
    });

    await createExcelFile(data, columns, fileName, title, headerColor);
  };

  const handleExportCustomers = async () => {
    if (filteredCustomers.length === 0) {
      showNotification('Không có dữ liệu khách hàng để xuất!', 'error');
      return;
    }
    const columns = [
      { header: 'Mã KH', key: 'code', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
      { header: 'Tên KH', key: 'name', width: 35 }
    ];
    await createExcelFile(filteredCustomers, columns, 'DanhSachKhachHang', 'DANH SÁCH KHÁCH HÀNG');
  };

  const handleProcessScanInput = async () => {
    if (!scanInput.trim()) return;

    const lines = scanInput.split('\n').filter(line => line.trim());
    const newHistoryEntries: LocationEntry[] = [];
    const inventoryAdditions: LocationEntry[] = [];
    const inventoryRemovals: string[] = [];
    const inventoryUpdates: LocationEntry[] = [];
    
    let activeLocation = currentLocation;
    let tempInventory = [...locationInventoryEntries];

    for (const line of lines) {
      const trimmedLine = line.trim();
      
      // Check if it's a location (starts with FB)
      if (trimmedLine.toUpperCase().startsWith('FB')) {
        activeLocation = trimmedLine;
        setCurrentLocation(activeLocation);
        continue;
      }

      // Otherwise process as QR code
      const qrcode = trimmedLine;
      const parsed = parseQRCode(qrcode);
      
      if (parsed) {
        // 1. Create History Entry
        const historyEntry: LocationEntry = {
          id: generateId(),
          qrcode,
          sku: parsed.sku,
          partner: parsed.partner,
          date: parsed.date || '',
          location: activeLocation,
          note: '',
          quantity: 1,
          type: 'input',
          scanType: scanMode,
          created_at: new Date().toISOString()
        };
        newHistoryEntries.push(historyEntry);

        // 2. Handle Inventory
        const existingIndex = tempInventory.findIndex(e => 
          e.qrcode === qrcode && (activeLocation ? e.location === activeLocation : true)
        );

        if (scanMode === 'INPUT') {
          if (existingIndex >= 0) {
            const existing = tempInventory[existingIndex];
            const updated = {
              ...existing,
              quantity: (existing.quantity || 0) + 1
            };
            tempInventory[existingIndex] = updated;
            inventoryUpdates.push(updated);
          } else {
            const newInv: LocationEntry = {
              ...historyEntry,
              id: generateId(),
              type: 'inventory'
            };
            tempInventory.push(newInv);
            inventoryAdditions.push(newInv);
          }
        } else {
          // OUTPUT: Remove from inventory
          if (existingIndex >= 0) {
            const existing = tempInventory[existingIndex];
            const newQty = (existing.quantity || 0) - 1;
            
            if (newQty <= 0) {
              inventoryRemovals.push(existing.id);
              tempInventory.splice(existingIndex, 1);
            } else {
              const updated = {
                ...existing,
                quantity: newQty
              };
              tempInventory[existingIndex] = updated;
              inventoryUpdates.push(updated);
            }
          }
        }
      }
    }

    if (newHistoryEntries.length > 0 || lines.some(l => l.trim().toUpperCase().startsWith('FB'))) {
      // Update local state
      if (newHistoryEntries.length > 0) {
        setLocationEntries(prev => [...prev, ...newHistoryEntries]);
        setLocationInventoryEntries(tempInventory);
      }

      try {
        if (newHistoryEntries.length > 0) {
          // Persist History
          await api.locationEntries.upsertAll(newHistoryEntries);
          
          // Persist Inventory Additions
          if (inventoryAdditions.length > 0) {
            await api.locationEntries.upsertAll(inventoryAdditions);
          }
          
          // Persist Inventory Updates
          if (inventoryUpdates.length > 0) {
            await api.locationEntries.upsertAll(inventoryUpdates);
          }
          
          // Persist Inventory Removals
          if (inventoryRemovals.length > 0) {
            await Promise.all(inventoryRemovals.map(id => api.locationEntries.delete(id)));
          }

          showNotification(`Đã xử lý ${newHistoryEntries.length} mã thành công!`);
        }
        setScanInput('');
      } catch (error) {
        console.error('Error processing scan entries:', error);
        showNotification('Lỗi khi xử lý dữ liệu scan.', 'error');
      }
    }
  };

  const handleExportLocations = async () => {
    const isInv = locationSubTab === 'inventory';
    const data = isInv ? filteredLocationInventoryEntries : filteredLocationEntries;
    
    if (data.length === 0) {
      showNotification(`Không có dữ liệu ${isInv ? 'tồn vị trí' : 'lịch sử vị trí'} để xuất!`, 'error');
      return;
    }

    const title = isInv ? 'BÁO CÁO TỒN KHO THEO VỊ TRÍ' : 'LỊCH SỬ NHẬP XUẤT VỊ TRÍ';
    const columns = [
      { header: 'QRCODE', key: 'qrcode', width: 25, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
      { header: 'Mã', key: 'sku', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
      { header: 'NCC', key: 'partner', width: 20 },
      { header: 'NGÀY', key: 'date', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
      { header: 'Vị trí', key: 'location', width: 15, alignment: { horizontal: 'center' as ExcelJS.Alignment['horizontal'] } },
    ];
    await createExcelFile(data, columns, isInv ? 'TonViTri' : 'LichSuViTri', title);
  };

  const handleAddProduct = async (e: React.FormEvent) => {
    e.preventDefault();
    let updatedProduct: Product;
    const isEditing = !!editingId;
    
    if (isEditing) {
      updatedProduct = { ...newProduct, id: editingId } as Product;
      setProducts(products.map(p => p.id === editingId ? updatedProduct : p));
      
      // Handle quantity adjustment if editing from inventory
      if (newProduct.quantity !== undefined) {
        const currentBatch = inventory.find(i => i.productId === editingId && i.lotNo === newProduct.lotNo && i.designationCode === newProduct.designationCode);
        const currentQty = currentBatch ? currentBatch.currentStock : 0;
        const diff = (newProduct.quantity || 0) - currentQty;
        
        if (diff !== 0) {
          const adjustment: Transaction = {
            id: generateId(),
            productId: updatedProduct.id,
            type: diff > 0 ? 'inbound' : 'outbound',
            quantity: Math.abs(diff),
            date: format(new Date(), 'dd/MM/yyyy'),
            partner: 'Điều chỉnh tồn kho',
            lotNo: newProduct.lotNo || '',
            ghiChu: 'Điều chỉnh số lượng từ giao diện tồn kho',
            designationCode: newProduct.designationCode || '',
            loaiChiDinh: newProduct.loaiChiDinh || ''
          };
          setTransactions(prev => [...prev, adjustment]);
          api.transactions.upsert(adjustment).catch(err => console.error('Error syncing adjustment:', err));
        }
      }
      
      setEditingId(null);
    } else {
      updatedProduct = {
        ...newProduct as Product,
        id: generateId(),
      };
      setProducts([...products, updatedProduct]);
      
      // If quantity is provided for a new product, create an initial inbound transaction
      if (newProduct.quantity && newProduct.quantity > 0) {
        const initialInbound: Transaction = {
          id: generateId(),
          productId: updatedProduct.id,
          type: 'inbound',
          quantity: newProduct.quantity,
          date: format(new Date(), 'dd/MM/yyyy'),
          partner: 'Khởi tạo tồn kho',
          lotNo: newProduct.lotNo || '',
          ghiChu: 'Số lượng ban đầu khi tạo sản phẩm',
          designationCode: newProduct.designationCode || '',
          loaiChiDinh: newProduct.loaiChiDinh || ''
        };
        setTransactions(prev => [...prev, initialInbound]);
        api.transactions.upsert(initialInbound).catch(err => console.error('Error syncing initial inbound:', err));
      }
    }
    
    try {
      await api.products.upsert(updatedProduct);
      showNotification('Sản phẩm đã được lưu.');
    } catch (error) {
      console.error('Error syncing product:', error);
      showNotification('Lỗi khi lưu sản phẩm.', 'error');
    }
    
    setIsProductModalOpen(false);
    setNewProduct({ sku: '', name: '', category: '', unit: '', minStock: 0, lotNo: '', ghiChu: '', designationCode: '', loaiChiDinh: '', quantity: 0 });
  };

  const handleAddTransaction = async (e: React.FormEvent) => {
    e.preventDefault();
    let updatedTransaction: Transaction;
    if (editingId) {
      updatedTransaction = { ...newTransaction, id: editingId } as Transaction;
      setTransactions(transactions.map(t => t.id === editingId ? updatedTransaction : t));
      setEditingId(null);
    } else {
      updatedTransaction = {
        ...newTransaction as Transaction,
        id: generateId(),
        type: activeTab === 'inbound' ? 'inbound' : 'outbound'
      };
      setTransactions([...transactions, updatedTransaction]);
    }
    
    try {
      await api.transactions.upsert(updatedTransaction);
      showNotification('Giao dịch đã được lưu.');
    } catch (error) {
      console.error('Error syncing transaction:', error);
      showNotification('Lỗi khi lưu giao dịch.', 'error');
    }
    
    setIsTransactionModalOpen(false);
    setNewTransaction({ productId: '', type: 'inbound', quantity: 0, date: format(new Date(), 'dd/MM/yyyy'), partner: '', loaiChiDinh: '', lotNo: '', ghiChu: '', designationCode: '' });
  };

  const handleAddCustomer = async (e: React.FormEvent) => {
    e.preventDefault();
    let updatedCustomer: Customer;
    if (editingId) {
      updatedCustomer = { ...newCustomer, id: editingId } as Customer;
      setCustomers(customers.map(c => c.id === editingId ? updatedCustomer : c));
      setEditingId(null);
    } else {
      updatedCustomer = {
        ...newCustomer as Customer,
        id: generateId(),
      };
      setCustomers([...customers, updatedCustomer]);
    }
    
    try {
      await api.customers.upsert(updatedCustomer);
      showNotification('Khách hàng đã được lưu.');
    } catch (error) {
      console.error('Error syncing customer:', error);
      showNotification('Lỗi khi lưu khách hàng.', 'error');
    }
    
    setIsCustomerModalOpen(false);
    setNewCustomer({ code: '', name: '' });
  };

  const handleLocationSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    let updatedEntry: LocationEntry;
    if (editingId) {
      const normalizedDate = normalizeDateForMatching(newLocationEntry.date || '');
      updatedEntry = { 
        ...newLocationEntry, 
        date: normalizedDate || newLocationEntry.date,
        id: editingId, 
        type: locationSubTab === 'inventory' ? 'inventory' : 'input' 
      } as LocationEntry;
      if (locationSubTab === 'input' || locationSubTab === 'output') {
        setLocationEntries(locationEntries.map(entry => entry.id === editingId ? updatedEntry : entry));
      } else {
        setLocationInventoryEntries(locationInventoryEntries.map(entry => entry.id === editingId ? updatedEntry : entry));
      }
      setEditingId(null);
    } else {
      const normalizedDate = normalizeDateForMatching(newLocationEntry.date || '');
      const finalDate = normalizedDate || newLocationEntry.date || '';

      if (locationSubTab === 'inventory') {
        const existingIndex = locationInventoryEntries.findIndex(
          entry => entry.qrcode === newLocationEntry.qrcode && entry.location === newLocationEntry.location
        );
        if (existingIndex >= 0) {
          const existing = locationInventoryEntries[existingIndex];
          updatedEntry = {
            ...existing,
            sku: newLocationEntry.sku || existing.sku,
            partner: newLocationEntry.partner || existing.partner,
            date: finalDate || existing.date,
            note: newLocationEntry.note || existing.note,
            quantity: (existing.quantity || 0) + (newLocationEntry.quantity || 0),
          };
          setLocationInventoryEntries(prev => prev.map((e, i) => i === existingIndex ? updatedEntry : e));
          setEditingId(null);
        } else {
          updatedEntry = {
            ...newLocationEntry as LocationEntry,
            date: finalDate,
            id: generateId(),
            type: locationSubTab === 'inventory' ? 'inventory' : 'input',
            created_at: new Date().toISOString()
          };
          setLocationInventoryEntries(prev => [...prev, updatedEntry]);
        }
      } else {
        updatedEntry = {
          ...newLocationEntry as LocationEntry,
          date: finalDate,
          id: generateId(),
          type: 'input',
          scanType: scanMode,
          created_at: new Date().toISOString()
        };
        setLocationEntries([updatedEntry, ...locationEntries]);

        const existingIndex = locationInventoryEntries.findIndex(
          entry => entry.qrcode === newLocationEntry.qrcode && entry.location === newLocationEntry.location
        );

        if (scanMode === 'INPUT') {
          if (existingIndex >= 0) {
            const existing = locationInventoryEntries[existingIndex];
            const updatedInventoryEntry: LocationEntry = {
              ...existing,
              sku: newLocationEntry.sku || existing.sku,
              partner: newLocationEntry.partner || existing.partner,
              date: newLocationEntry.date || existing.date,
              note: newLocationEntry.note || existing.note,
              quantity: (existing.quantity || 0) + (newLocationEntry.quantity || 0),
            };
            setLocationInventoryEntries(prev => prev.map((e, i) => i === existingIndex ? updatedInventoryEntry : e));
            try {
              await api.locationEntries.upsert(updatedInventoryEntry);
            } catch (error) {
              console.error('Error syncing inventory entry:', error);
            }
          } else {
            const newInventoryEntry: LocationEntry = {
              ...newLocationEntry as LocationEntry,
              id: generateId(),
              type: 'inventory',
              created_at: new Date().toISOString()
            };
            setLocationInventoryEntries(prev => [newInventoryEntry, ...prev]);
            try {
              await api.locationEntries.upsert(newInventoryEntry);
            } catch (error) {
              console.error('Error syncing inventory entry:', error);
            }
          }
        } else if (scanMode === 'OUTPUT') {
          if (existingIndex >= 0) {
            const existing = locationInventoryEntries[existingIndex];
            const newQty = (existing.quantity || 0) - (newLocationEntry.quantity || 0);
            
            if (newQty <= 0) {
              setLocationInventoryEntries(prev => prev.filter((_, i) => i !== existingIndex));
              try {
                await api.locationEntries.delete(existing.id);
              } catch (error) {
                console.error('Error deleting inventory entry:', error);
              }
            } else {
              const updatedInventoryEntry: LocationEntry = {
                ...existing,
                quantity: newQty
              };
              setLocationInventoryEntries(prev => prev.map((e, i) => i === existingIndex ? updatedInventoryEntry : e));
              try {
                await api.locationEntries.upsert(updatedInventoryEntry);
              } catch (error) {
                console.error('Error updating inventory entry:', error);
              }
            }
          }
        }
      }
    }
    
    try {
      await api.locationEntries.upsert(updatedEntry);
      showNotification('Vị trí đã được lưu.');
    } catch (error) {
      console.error('Error syncing location entry:', error);
      showNotification('Lỗi khi lưu vị trí.', 'error');
    }
    
    setIsLocationModalOpen(false);
    setNewLocationEntry({ qrcode: '', sku: '', partner: '', date: '', location: '', note: '' });
  };

  const processLocationData = async (data: any[]) => {
    const newEntries: any[] = data.map((row: any) => {
      const normalizedRow: any = {};
      Object.keys(row).forEach(key => {
        normalizedRow[key.toLowerCase().trim()] = row[key];
      });

      const qrcode = String(normalizedRow['qrcode'] || normalizedRow['qr code'] || row['QR Code'] || row['qrcode'] || '').trim();
      const parsed = parseQRCode(qrcode);
      
      const rawDate = String(normalizedRow['date'] || normalizedRow['ngày'] || row['Ngày'] || row['date'] || '');
      const normalizedDate = normalizeDateForMatching(rawDate);
      const finalDate = normalizedDate || rawDate || parsed?.date || '';

      return {
        qrcode,
        sku: normalizedRow['sku'] ?? row['SKU'] ?? row['sku'] ?? parsed?.sku ?? '',
        partner: normalizedRow['partner'] ?? normalizedRow['đối tác'] ?? row['Đối tác'] ?? row['partner'] ?? parsed?.partner ?? '',
        date: finalDate,
        location: normalizedRow['location'] ?? normalizedRow['vị trí'] ?? normalizedRow['vi tri'] ?? row['Vị trí'] ?? row['location'] ?? '',
        note: normalizedRow['note'] ?? normalizedRow['ghi chú'] ?? normalizedRow['cuộn'] ?? normalizedRow['cuon'] ?? row['Cuộn'] ?? row['Ghi chú'] ?? row['note'] ?? '',
        quantity: parseInt(String(normalizedRow['quantity'] ?? normalizedRow['số lượng'] ?? normalizedRow['so luong'] ?? row['Số lượng'] ?? row['quantity'] ?? '1')) || 1
      };
    }).filter(entry => entry.qrcode);

    if (newEntries.length === 0) return;

    const currentScanType = locationSubTab === 'output' ? 'OUTPUT' : 'INPUT';
    const logEntriesToAdd: LocationEntry[] = newEntries.map(entry => ({
      ...entry,
      id: generateId(),
      type: 'input' as const,
      scanType: currentScanType
    }));

    setLocationEntries(prev => [...logEntriesToAdd, ...prev]);

    let finalInventoryEntries: LocationEntry[] = [];
    let entriesToDeleteFromDb: LocationEntry[] = [];
    
    let prevEntriesForSync: LocationEntry[] = [];
    setLocationInventoryEntries(prev => {
      prevEntriesForSync = prev;
      const inventoryGrouped = new Map<string, LocationEntry>();
      prev.forEach(entry => {
        const key = `${entry.qrcode}|${entry.location || ''}`;
        inventoryGrouped.set(key, { ...entry });
      });

      newEntries.forEach(entry => {
        if (entry.qrcode) {
          const key = `${entry.qrcode}|${entry.location || ''}`;
          if (currentScanType === 'INPUT') {
            if (inventoryGrouped.has(key)) {
              const existing = inventoryGrouped.get(key)!;
              existing.sku = entry.sku !== undefined ? entry.sku : existing.sku;
              existing.partner = entry.partner !== undefined ? entry.partner : existing.partner;
              existing.date = entry.date !== undefined ? entry.date : existing.date;
              existing.note = entry.note !== undefined ? entry.note : existing.note;
              existing.quantity = (existing.quantity || 0) + (entry.quantity || 0);
            } else {
              inventoryGrouped.set(key, { 
                ...entry, 
                id: generateId(), 
                type: 'inventory' 
              });
            }
          } else if (currentScanType === 'OUTPUT') {
            if (inventoryGrouped.has(key)) {
              const existing = inventoryGrouped.get(key)!;
              existing.quantity = (existing.quantity || 0) - (entry.quantity || 0);
              if (existing.quantity <= 0) {
                inventoryGrouped.delete(key);
              }
            }
          }
        }
      });

      const newFinal = Array.from(inventoryGrouped.values());
      
      const prevIds = new Set(prev.map(e => e.id));
      const newIds = new Set(newFinal.map(e => e.id));
      entriesToDeleteFromDb = prev.filter(e => prevIds.has(e.id) && !newIds.has(e.id));
      
      finalInventoryEntries = newFinal;
      return newFinal;
    });
    
    try {
      const syncTasks: Promise<void>[] = [
        api.locationEntries.upsertAll(logEntriesToAdd),
      ];
      
      // Only upsert changed inventory entries
      const changedInventoryEntries = finalInventoryEntries.filter(e => {
        const prev = prevEntriesForSync.find(p => p.id === e.id);
        if (!prev) return true; // New
        return JSON.stringify(prev) !== JSON.stringify(e); // Changed
      });

      if (changedInventoryEntries.length > 0) {
        syncTasks.push(api.locationEntries.upsertAll(changedInventoryEntries));
      }

      if (entriesToDeleteFromDb.length > 0) {
        syncTasks.push(api.locationEntries.deleteMany(entriesToDeleteFromDb.map(e => e.id)));
      }
      await Promise.all(syncTasks);
      showNotification('Nhập file vị trí thành công.');
    } catch (error) {
      console.error('Error syncing location entries:', error);
      showNotification('Lỗi khi lưu dữ liệu vị trí.', 'error');
    }
  };


  const handleDelete = async (id: string, type: 'product' | 'transaction' | 'customer' | 'location' | 'savedDeliveryNote') => {
    if (type === 'product') {
      setProducts(prev => prev.filter(p => p.id !== id));
      setTransactions(prev => prev.filter(t => t.productId !== id));
    } else if (type === 'transaction') {
      setTransactions(prev => prev.filter(t => t.id !== id));
    } else if (type === 'customer') {
      setCustomers(prev => prev.filter(c => c.id !== id));
    } else if (type === 'location') {
      const qrcode = deleteTarget?.qrcode;
      if (qrcode) {
        setLocationEntries(prev => prev.filter(e => e.qrcode !== qrcode));
        setLocationInventoryEntries(prev => prev.filter(e => e.qrcode !== qrcode));
      } else {
        setLocationEntries(prev => prev.filter(e => e.id !== id));
        setLocationInventoryEntries(prev => prev.filter(e => e.id !== id));
      }
    } else if (type === 'savedDeliveryNote') {
      setSavedDeliveryNotes(prev => prev.filter(n => n.id !== id));
    }

    try {
      if (type === 'product') {
        await api.transactions.deleteByProductId(id);
        await api.products.delete(id);
      } else if (type === 'transaction') {
        await api.transactions.delete(id);
      } else if (type === 'customer') {
        await api.customers.delete(id);
      } else if (type === 'location') {
        const qrcode = deleteTarget?.qrcode;
        if (qrcode) {
          await api.locationEntries.deleteByQRCode(qrcode);
        } else {
          await api.locationEntries.delete(id);
        }
      } else if (type === 'savedDeliveryNote') {
        await api.savedDeliveryNotes.delete(id);
      }
      showNotification('Đã xóa dữ liệu thành công.');
    } catch (error: any) {
      console.error('Error syncing deletion:', error);
      showNotification('Lỗi khi xóa dữ liệu: ' + (error.message || 'Vui lòng thử lại sau.'), 'error');
      loadData();
    }
  };

  const handleBulkDelete = async () => {
    const idsToDelete = [...selectedRows];
    const currentTab = activeTab;
    const subTab = locationSubTab;

    if (currentTab === 'inventory') {
      const inventoryMap = new Map<string, InventoryItem>(inventory.map(i => [i.id, i]));
      const productIds = idsToDelete.map(batchKey => {
        const item = inventoryMap.get(batchKey);
        return item?.productId;
      }).filter(Boolean) as string[];
      const uniqueProductIds = [...new Set(productIds)];
      
      if (uniqueProductIds.length === 0) {
        setSelectedRows([]);
        setIsDeleteConfirmOpen(false);
        return;
      }

      setProducts(prev => prev.filter(p => !uniqueProductIds.includes(p.id)));
      setTransactions(prev => prev.filter(t => !uniqueProductIds.includes(t.productId)));
      
      try {
        await api.transactions.deleteByProductIds(uniqueProductIds);
        await api.products.deleteMany(uniqueProductIds);
        showNotification(`Đã xóa ${uniqueProductIds.length} mặt hàng thành công.`);
      } catch (error: any) {
        console.error('Error in bulk delete inventory:', error);
        showNotification('Lỗi khi xóa hàng loạt tồn kho: ' + (error.message || 'Vui lòng kiểm tra kết nối mạng.'), 'error');
        loadData();
      }
    } else if (currentTab === 'inbound' || currentTab === 'outbound') {
      setTransactions(prev => prev.filter(t => !idsToDelete.includes(t.id)));
      try {
        await api.transactions.deleteMany(idsToDelete);
        showNotification('Đã xóa các giao dịch thành công.');
      } catch (error) {
        showNotification('Lỗi khi xóa giao dịch.', 'error');
        loadData();
      }
    } else if (currentTab === 'customers') {
      setCustomers(prev => prev.filter(c => !idsToDelete.includes(c.id)));
      try {
        await api.customers.deleteMany(idsToDelete);
        showNotification('Đã xóa các khách hàng thành công.');
      } catch (error) {
        showNotification('Lỗi khi xóa khách hàng.', 'error');
        loadData();
      }
    } else if (currentTab === 'location') {
      if (subTab === 'input' || subTab === 'output') {
        setLocationEntries(prev => prev.filter(e => !idsToDelete.includes(e.id)));
        try {
          await api.locationEntries.deleteMany(idsToDelete);
          showNotification('Đã xóa các mục vị trí thành công.');
        } catch (error) {
          showNotification('Lỗi khi xóa vị trí.', 'error');
          loadData();
        }
      } else {
        const qrcodes = idsToDelete.map(id => {
          const entry = locationInventoryEntries.find(e => e.id === id);
          return entry?.qrcode;
        }).filter(Boolean) as string[];
        const uniqueQrcodes = [...new Set(qrcodes)];
        
        setLocationEntries(prev => prev.filter(e => !uniqueQrcodes.includes(e.qrcode)));
        setLocationInventoryEntries(prev => prev.filter(e => !uniqueQrcodes.includes(e.qrcode)));
        try {
          await api.locationEntries.deleteByQRCodes(uniqueQrcodes);
          showNotification('Đã xóa các mục tồn vị trí thành công.');
        } catch (error) {
          showNotification('Lỗi khi xóa tồn vị trí.', 'error');
          loadData();
        }
      }
    } else if (currentTab === 'deliveryNote') {
      if (deliveryNoteSubTab === 'preview') {
        setDeliveryNotes(prev => {
          const filtered = prev.filter(item => !idsToDelete.includes(item.id));
          const updated = recalculateActualQty(filtered);
          api.deliveryNotes.upsertAll(updated).catch(err => console.error('Error syncing delivery notes:', err));
          return updated;
        });
        showNotification('Đã xóa các mục phiếu giao nhận thành công.');
      } else {
        setSavedDeliveryNotes(prev => prev.filter(n => !idsToDelete.includes(n.id)));
        try {
          await api.savedDeliveryNotes.deleteMany(idsToDelete);
          showNotification('Đã xóa các phiếu lưu thành công.');
        } catch (error) {
          console.error('Error in bulk delete saved delivery notes:', error);
          showNotification('Lỗi khi xóa phiếu lưu.', 'error');
          loadData();
        }
      }
    }

    setSelectedRows([]);
  };

  const handleWipeAllData = async () => {
    try {
      if (isSupabaseConfigured) {
        await api.products.deleteAll();
        await api.transactions.deleteAll();
        await api.locationEntries.deleteAll();
      }
      setProducts([]);
      setTransactions([]);
      setLocationEntries([]);
      setSelectedRows([]);
      showNotification('Đã xóa toàn bộ dữ liệu Tồn kho, Nhập/Xuất & Vị trí', 'success');
    } catch (error) {
      console.error('Lỗi khi xóa đồng loạt:', error);
      showNotification('Có lỗi xảy ra khi xóa toàn bộ dữ liệu', 'error');
    }
  };

  const confirmDelete = async () => {
    if (!deleteTarget) return;
    const target = { ...deleteTarget };
    setIsDeleteConfirmOpen(false);
    setDeleteTarget(null);

    if (target.id === 'bulk') {
      await handleBulkDelete();
    } else if (target.id === 'wipe_all_data') {
      await handleWipeAllData();
    } else {
      await handleDelete(target.id, target.type as any);
    }
  };

  const toggleRowSelection = (id: string) => {
    setSelectedRows(prev => prev.includes(id) ? prev.filter(rowId => rowId !== id) : [...prev, id]);
  };

  const inventory = useMemo(() => {
    const batches: Record<string, InventoryItem> = {};
    const productMap = new Map<string, Product>(products.map(p => [p.id, p]));
    const productsWithTransactions = new Set<string>();

    transactions.forEach(t => {
      const product = productMap.get(t.productId);
      if (!product) return;

      productsWithTransactions.add(t.productId);

      const batchKey = `${t.productId}-${t.lotNo || ''}-${t.designationCode || ''}`;
      
      if (!batches[batchKey]) {
        batches[batchKey] = {
          ...product,
          id: batchKey,
          productId: product.id,
          lotNo: t.lotNo || '',
          ghiChu: t.ghiChu || '',
          designationCode: t.designationCode || '',
          loaiChiDinh: t.loaiChiDinh || '',
          rpro: t.rpro || product.rpro || '',
          totalInbound: 0,
          totalOutbound: 0,
          currentStock: 0,
          inboundDate: t.type === 'inbound' ? extractDateFromLot(t.lotNo || '') : 9999999999999
        };
      } else {
        if (t.loaiChiDinh && !batches[batchKey].loaiChiDinh?.includes(t.loaiChiDinh)) {
          // Merge loaiChiDinh for reference if it's different
          batches[batchKey].loaiChiDinh = batches[batchKey].loaiChiDinh 
            ? `${batches[batchKey].loaiChiDinh}, ${t.loaiChiDinh}` 
            : t.loaiChiDinh;
        }
        if (t.type === 'inbound') {
          const tDate = extractDateFromLot(t.lotNo || '');
          if (tDate < (batches[batchKey].inboundDate || 9999999999999)) {
            batches[batchKey].inboundDate = tDate;
          }
        }
      }

      if (t.type === 'inbound') {
        batches[batchKey].totalInbound += t.quantity;
      } else {
        batches[batchKey].totalOutbound += t.quantity;
      }
      batches[batchKey].currentStock = batches[batchKey].totalInbound - batches[batchKey].totalOutbound;
    });

    products.forEach(product => {
      if (!productsWithTransactions.has(product.id)) {
        const batchKey = `${product.id}-empty`;
        batches[batchKey] = {
          ...product,
          id: batchKey,
          productId: product.id,
          totalInbound: 0,
          totalOutbound: 0,
          currentStock: 0,
          inboundDate: 9999999999999
        };
      }
    });

    return Object.values(batches);
  }, [products, transactions, extractDateFromLot]);

  const getStockDetailByDesignation = useCallback((sku: string) => {
    const skuInventory = inventory.filter(item => item.sku.toLowerCase().trim() === sku.toLowerCase().trim());
    if (skuInventory.length === 0) return 'Không có tồn';
    
    const summary = new Map<string, number>();
    skuInventory.forEach(item => {
      const type = item.loaiChiDinh || 'Chung';
      summary.set(type, (summary.get(type) || 0) + item.currentStock);
    });
    
    return Array.from(summary.entries())
      .map(([type, qty]) => `${type}: ${qty.toLocaleString()}`)
      .join(' | ');
  }, [inventory]);

  const analyzeGroupStock = useCallback((sku: string, groupItems: DeliveryNoteItem[]) => {
    // Get all inventory for this SKU and create a working copy of stock
    const skuInventory = inventory
      .filter(item => item.sku.toLowerCase().trim() === sku.toLowerCase().trim())
      .map(item => ({ 
        ...item, 
        tempStock: item.currentStock,
        normalizedCodes: (item.designationCode || '').toLowerCase().replace(/\s+/g, '').split('/')
      }));

    if (skuInventory.length === 0) {
      return { 
        detail: 'Không có tồn', 
        shortage: groupItems.reduce((sum, gi) => sum + (gi.actualQty || gi.qtyErp), 0) 
      };
    }

    let totalShortage = 0;

    groupItems.forEach(gi => {
      let needed = gi.actualQty || gi.qtyErp;
      const targetRpro = (gi.ovnProductionOrder || '').toLowerCase().trim();
      const targetSo = (gi.ovnSaleOrder || '').toLowerCase().trim();
      const targetNo = (gi.noCode || '').toLowerCase().trim();

      // Helper to check if an inventory item matches any of the targets
      const isMatch = (invItem: any) => {
        return (targetRpro && invItem.normalizedCodes.includes(targetRpro)) || 
               (targetSo && invItem.normalizedCodes.includes(targetSo)) || 
               (targetNo && invItem.normalizedCodes.includes(targetNo));
      };

      // Priority 1: Specific Designation Match
      skuInventory.forEach(invItem => {
        if (needed <= 0) return;
        if (invItem.tempStock <= 0) return;
        if (isMatch(invItem)) {
          const take = Math.min(invItem.tempStock, needed);
          invItem.tempStock -= take;
          needed -= take;
        }
      });

      // Priority 2: General Stock (Empty Designation)
      if (needed > 0) {
        skuInventory.forEach(invItem => {
          if (needed <= 0) return;
          if (invItem.tempStock <= 0) return;
          if (!(invItem.designationCode || '').trim()) {
            const take = Math.min(invItem.tempStock, needed);
            invItem.tempStock -= take;
            needed -= take;
          }
        });
      }

      totalShortage += needed;
    });

    const detail = getStockDetailByDesignation(sku);
    return { detail, shortage: totalShortage };
  }, [inventory, getStockDetailByDesignation]);

  const getLocationByItemAndLot = useCallback((sku: string, lotNo: string) => {
    if (!sku) return 'Chưa có vị trí';
    
    const skuLower = sku.toLowerCase().trim();
    const normalizedLotDate = normalizeDateForMatching(lotNo || '');
    const rawLot = (lotNo || '').trim();
    
    const matches = locationInventoryEntries.filter(e => {
      const entrySku = (e.sku || '').toLowerCase().trim();
      if (entrySku !== skuLower) return false;

      const entryDate = (e.date || '').trim();
      const normalizedEntryDate = normalizeDateForMatching(entryDate);

      return entryDate === rawLot || 
             entryDate === normalizedLotDate || 
             normalizedEntryDate === rawLot || 
             normalizedEntryDate === normalizedLotDate;
    });
    
    if (matches.length === 0) return 'Chưa có vị trí';

    // Aggregate unique locations
    const locations = Array.from(new Set(matches.map(m => m.location).filter(Boolean)));
    return locations.length > 0 ? locations.join(', ') : 'Chưa có vị trí';
  }, [locationInventoryEntries, normalizeDateForMatching]);

  const filteredInventory = useMemo(() => {
    const query = searchQuery.toLowerCase();
    return inventory.filter(item => 
      item.name.toLowerCase().includes(query) ||
      item.sku.toLowerCase().includes(query)
    );
  }, [inventory, searchQuery]);

  // Tự động điền LOT NO khi người dùng nhập AWB/RPRO/NO trong Modal chỉnh sửa
  useEffect(() => {
    if (isDeliveryNoteEditModalOpen && tempDeliveryNoteItem.item) {
      const matches = findAssignedFabricLot(
        { 
          sku: tempDeliveryNoteItem.item, 
          rpro: tempDeliveryNoteItem.ovnProductionOrder || '', 
          no: tempDeliveryNoteItem.noCode || '' 
        },
        inventory
      );
      
      if (matches.length > 0) {
        const bestMatch = matches[0];
        // Tự động điền nếu Lot No đang trống
        if (!tempDeliveryNoteItem.lotNo) {
          const normalizedLotDate = normalizeDateForMatching(bestMatch.lotNo || '');
          const locationEntry = locationInventoryEntries.find(e => 
            e.sku === tempDeliveryNoteItem.item && e.date === normalizedLotDate
          );

          setTempDeliveryNoteItem(prev => ({
            ...prev,
            lotNo: bestMatch.lotNo,
            location: locationEntry ? locationEntry.location : 'Chưa có vị trí',
            stock: `${bestMatch.currentStock} (${bestMatch.loaiChiDinh || 'N/A'})`
          }));
        }
      } else if (tempDeliveryNoteItem.item && (tempDeliveryNoteItem.ovnProductionOrder || tempDeliveryNoteItem.noCode)) {
        // Nếu không tìm thấy và đã nhập đủ thông tin thì có thể xóa lotNo hoặc để người dùng tự nhập
        // Ở đây ta giữ nguyên để người dùng tự xử lý nếu muốn
      }
    }
  }, [
    tempDeliveryNoteItem.item, 
    tempDeliveryNoteItem.ovnProductionOrder, 
    tempDeliveryNoteItem.noCode, 
    isDeliveryNoteEditModalOpen,
    findAssignedFabricLot,
    inventory,
    locationInventoryEntries,
    normalizeDateForMatching
  ]);

  const autoAssignLotsFromUI = useCallback(() => {
    if (deliveryNotes.length === 0) {
      showNotification('Không có dữ liệu để xử lý', 'error');
      return;
    }

    // Luôn làm tròn lại số lượng thực tế trước khi gán Lot
    const roundedNotes = recalculateActualQty(deliveryNotes);

    const workingInventory = inventory.map(item => ({
      ...item,
      tempStock: item.currentStock
    }));

    const updatedNotes = roundedNotes.map(item => {
      const matches = findAssignedFabricLot(
        { sku: item.item, rpro: item.ovnProductionOrder, no: item.noCode },
        workingInventory
      );

      if (matches.length === 0) {
        return {
          ...item,
          assignedLots: [],
          lotNo: '',
          actualQty: 0,
          actualIssuedQty: 0,
          location: '',
          stock: 'Không tìm thấy tồn kho phù hợp theo chỉ định'
        };
      }

      let remainingNeeded = (item.actualQty !== undefined && item.actualQty !== null && item.actualQty !== 0) ? item.actualQty : (item.actualQty === 0 ? 0 : item.qtyErp);
      // Simplified: if actualQty is set (even if 0), use it. Otherwise use qtyErp.
      const targetQty = (item.actualQty !== undefined && item.actualQty !== null) ? item.actualQty : item.qtyErp;
      remainingNeeded = targetQty;
      const assignedLots = [];

      for (const match of matches) {
        if (remainingNeeded <= 0) break;

        const available = match.tempStock !== undefined ? match.tempStock : match.currentStock;
        if (available <= 0) continue;

        const allocation = Math.min(available, remainingNeeded);
        
        if (match.tempStock !== undefined) {
          match.tempStock -= allocation;
        } else {
          match.tempStock = match.currentStock - allocation;
        }

        const location = getLocationByItemAndLot(item.item, match.lotNo || '');

        assignedLots.push({
          lotNo: match.lotNo || '',
          qty: allocation,
          stock: `${match.currentStock} (${match.loaiChiDinh || 'N/A'})`,
          location: location,
          remark: match.designationCode || '',
          loaiChiDinh: match.loaiChiDinh || ''
        });

        remainingNeeded -= allocation;
      }

      const totalQty = assignedLots.reduce((sum, l) => sum + l.qty, 0);

      return {
        ...item,
        assignedLots,
        lotNo: assignedLots.map(l => l.lotNo).join(', '),
        // Keep actualQty as the rounded target, update actualIssuedQty with assigned amount
        actualIssuedQty: totalQty,
        location: assignedLots.map(l => l.location).join(', '),
        stock: remainingNeeded > 0 ? `Thiếu ${remainingNeeded.toLocaleString()}` : `Đã gán ${assignedLots.length} Lot`
      };
    });

    setDeliveryNotes(updatedNotes);
    showNotification('Đã tự động gán Lot No (gộp thông tin vào một dòng)', 'success');
  }, [deliveryNotes, inventory, findAssignedFabricLot, locationInventoryEntries, normalizeDateForMatching]);

  const filteredCustomers = useMemo(() => {
    const query = searchQuery.toLowerCase();
    return customers.filter(c => 
      c.name.toLowerCase().includes(query) ||
      c.code.toLowerCase().includes(query)
    );
  }, [customers, searchQuery]);

  const filteredLocationEntries = useMemo(() => {
    const query = locationSearch.toLowerCase();
    const currentScanType = locationSubTab === 'input' ? 'INPUT' : 'OUTPUT';
    
    return locationEntries.filter(entry => {
      // Filter out location entries that are actually location codes (start with FB)
      if (entry.qrcode && entry.qrcode.toUpperCase().startsWith('FB')) return false;

      const matchesSearch = (entry.sku || '').toLowerCase().includes(query) ||
        (entry.partner || '').toLowerCase().includes(query) ||
        (entry.location || '').toLowerCase().includes(query) ||
        (entry.qrcode || '').toLowerCase().includes(query) ||
        (entry.note || '').toLowerCase().includes(query);
      
      if (locationSubTab === 'inventory') return matchesSearch;
      return matchesSearch && entry.scanType === currentScanType;
    });
  }, [locationEntries, locationSearch, locationSubTab]);

  const filteredLocationInventoryEntries = useMemo(() => {
    const query = locationSearch.toLowerCase();
    
    const groupedMap = new Map<string, LocationEntry>();
    
    locationInventoryEntries.forEach(entry => {
      const qrcode = entry.qrcode;
      // Filter out location entries that are actually location codes (start with FB)
      if (qrcode && qrcode.toUpperCase().startsWith('FB')) return;

      if (groupedMap.has(qrcode)) {
        const existing = groupedMap.get(qrcode)!;
        
        const currentLocs = (existing.location || '').split(',').map(l => l.trim()).filter(Boolean);
        const newLoc = (entry.location || '').trim();
        if (newLoc && !currentLocs.includes(newLoc)) {
          existing.location = currentLocs.length > 0 ? `${existing.location}, ${newLoc}` : newLoc;
        }
        
        const currentNotes = (existing.note || '').split(',').map(n => n.trim()).filter(Boolean);
        const newNote = (entry.note || '').trim();
        if (newNote && !currentNotes.includes(newNote)) {
          existing.note = currentNotes.length > 0 ? `${existing.note}, ${newNote}` : newNote;
        }

        existing.quantity = (existing.quantity || 0) + (entry.quantity || 0);
      } else {
        groupedMap.set(qrcode, { ...entry });
      }
    });

    const groupedEntries = Array.from(groupedMap.values());

    return groupedEntries.filter(entry => 
      (entry.sku || '').toLowerCase().includes(query) ||
      (entry.partner || '').toLowerCase().includes(query) ||
      (entry.location || '').toLowerCase().includes(query) ||
      (entry.qrcode || '').toLowerCase().includes(query) ||
      (entry.note || '').toLowerCase().includes(query)
    );
  }, [locationInventoryEntries, locationSearch]);

  const parseDate = useCallback((dateStr: string) => {
    let parsed = parse(dateStr, 'dd/MM/yyyy', new Date());
    if (isValid(parsed)) return parsed;
    
    parsed = new Date(dateStr);
    if (isValid(parsed)) return parsed;

    return new Date(0);
  }, []);

  const filteredTransactions = useMemo(() => {
    const productMap = new Map<string, Product>(products.map(p => [p.id, p]));
    const query = searchQuery.toLowerCase();
    
    const transactionWithTimestamps = transactions
      .filter(t => {
        const product = productMap.get(t.productId);
        return product?.name.toLowerCase().includes(query) ||
               product?.sku.toLowerCase().includes(query) ||
               t.partner.toLowerCase().includes(query);
      })
      .map(t => ({ ...t, timestamp: parseDate(t.date).getTime() }));

    return transactionWithTimestamps.sort((a, b) => b.timestamp - a.timestamp);
  }, [transactions, products, searchQuery, parseDate]);

  const filteredDeliveryNotes = useMemo(() => {
    const query = searchQuery.toLowerCase();
    return deliveryNotes.filter(item => 
      item.item.toLowerCase().includes(query) ||
      item.materialName.toLowerCase().includes(query) ||
      item.ovnSaleOrder.toLowerCase().includes(query) ||
      item.ovnProductionOrder.toLowerCase().includes(query) ||
      item.lotNo.toLowerCase().includes(query) ||
      item.customerCode.toLowerCase().includes(query)
    );
  }, [deliveryNotes, searchQuery]);

  const processData = (data: any[]) => {
    const normalizeKey = (key: string) => {
      return key.toLowerCase().trim()
        .replace(/\s+/g, ''); // Remove all spaces, keep dots
    };

    if (activeTab === 'inventory') {
      setProducts(prev => {
        const updatedProducts = [...prev];
        const skuToProductIndex = new Map(updatedProducts.map((p, i) => [p.sku, i]));
        const productsToUpsert: Product[] = [];

        data.forEach((row) => {
          const normalizedRow: any = {};
          Object.keys(row).forEach(key => {
            normalizedRow[normalizeKey(key)] = row[key];
          });

          let sku = String(
            normalizedRow['sku'] ?? 
            normalizedRow['mãhàng'] ?? 
            normalizedRow['mãhànghóa'] ??
            normalizedRow['mãhh'] ??
            normalizedRow['mãsp'] ?? 
            normalizedRow['mãsảnphẩm'] ?? 
            normalizedRow['mãvậttư'] ??
            normalizedRow['itemno'] ??
            normalizedRow['itemcode'] ??
            normalizedRow['mã'] ??
            normalizedRow['item'] ??
            row.sku ?? row.SKU ?? row['Mã Hàng'] ?? ''
          ).trim();

          if (!sku) return;

          let name = String(
            normalizedRow['name'] ?? 
            normalizedRow['tênhàng'] ?? 
            normalizedRow['tênhànghóa'] ??
            normalizedRow['tênhh'] ??
            normalizedRow['tênsp'] ?? 
            normalizedRow['tênsảnphẩm'] ?? 
            normalizedRow['tênvậttư'] ??
            normalizedRow['productname'] ?? 
            normalizedRow['itemname'] ??
            normalizedRow['description'] ??
            normalizedRow['môtả'] ??
            normalizedRow['tên'] ??
            row.name ?? row.Name ?? row['Tên hàng'] ?? ''
          ).trim();

          const rowLot = normalizedRow['lotno'] !== undefined ? String(normalizedRow['lotno']).trim() : 
                        (normalizedRow['lotno.'] !== undefined ? String(normalizedRow['lotno.']).trim() : undefined);
          const rowGhiChu = normalizedRow['ghichu'] !== undefined ? String(normalizedRow['ghichu']).trim() : 
                           (normalizedRow['ghichú'] !== undefined ? String(normalizedRow['ghichú']).trim() : undefined);
          const rowDesignation = normalizedRow['designationcode'] !== undefined ? String(normalizedRow['designationcode']).replace(/\s+/g, '') : 
                                (normalizedRow['mãchỉđịnh'] !== undefined ? String(normalizedRow['mãchỉđịnh']).replace(/\s+/g, '') : undefined);
          const rowLoai = normalizedRow['loaichidinh'] !== undefined ? String(normalizedRow['loaichidinh']).trim() : 
                         (normalizedRow['loạichỉđịnh'] !== undefined ? String(normalizedRow['loạichỉđịnh']).trim() : undefined);

          const existingIndex = skuToProductIndex.get(sku);
          const productData: Product = {
            id: existingIndex !== undefined ? updatedProducts[existingIndex].id : generateId(),
            sku: sku,
            name: name || (existingIndex !== undefined ? updatedProducts[existingIndex].name : 'Sản phẩm mới'),
            category: normalizedRow['category'] !== undefined ? String(normalizedRow['category']).trim() : 
                      (normalizedRow['loại'] !== undefined ? String(normalizedRow['loại']).trim() : 
                      (existingIndex !== undefined ? updatedProducts[existingIndex].category : 'Chưa phân loại')),
            unit: normalizedRow['unit'] !== undefined ? String(normalizedRow['unit']).trim() : 
                  (normalizedRow['đơnvị'] !== undefined ? String(normalizedRow['đơnvị']).trim() : 
                  (existingIndex !== undefined ? updatedProducts[existingIndex].unit : 'Cái')),
            minStock: normalizedRow['minstock'] !== undefined ? Number(normalizedRow['minstock']) : 
                      (normalizedRow['tồntốithiểu'] !== undefined ? Number(normalizedRow['tồntốithiểu']) : 
                      (existingIndex !== undefined ? updatedProducts[existingIndex].minStock : 0)),
            lotNo: rowLot ?? (existingIndex !== undefined ? updatedProducts[existingIndex].lotNo : ''),
            ghiChu: rowGhiChu ?? (existingIndex !== undefined ? updatedProducts[existingIndex].ghiChu : ''),
            designationCode: rowDesignation ?? (existingIndex !== undefined ? updatedProducts[existingIndex].designationCode : ''),
            loaiChiDinh: rowLoai ?? (existingIndex !== undefined ? updatedProducts[existingIndex].loaiChiDinh : '')
          };

          if (existingIndex !== undefined) {
            // Only update if something changed
            const existing = updatedProducts[existingIndex];
            const hasChanged = 
              existing.name !== productData.name ||
              existing.category !== productData.category ||
              existing.unit !== productData.unit ||
              existing.minStock !== productData.minStock ||
              existing.lotNo !== productData.lotNo ||
              existing.ghiChu !== productData.ghiChu ||
              existing.designationCode !== productData.designationCode ||
              existing.loaiChiDinh !== productData.loaiChiDinh;

            if (hasChanged) {
              updatedProducts[existingIndex] = productData;
              productsToUpsert.push(productData);
            }
          } else {
            skuToProductIndex.set(sku, updatedProducts.length);
            updatedProducts.push(productData);
            productsToUpsert.push(productData);
          }
        });

        if (productsToUpsert.length > 0) {
          api.products.upsertAll(productsToUpsert).catch(err => console.error('Error syncing products:', err));
        }
        return updatedProducts;
      });
    } else if (activeTab === 'inbound' || activeTab === 'outbound') {
      setProducts(prevProducts => {
        const updatedProducts = [...prevProducts];
        const newProductsToAdd: Product[] = [];
        const changedProducts: Product[] = [];
        const skuToProductIndex = new Map(updatedProducts.map((p, i) => [p.sku, i]));
        const skuToNewProductIndex = new Map<string, number>();

        data.forEach((row) => {
          const normalizedRow: any = {};
          Object.keys(row).forEach(key => {
            normalizedRow[normalizeKey(key)] = row[key];
          });

          let sku = String(
            normalizedRow['sku'] ?? 
            normalizedRow['mãhàng'] ?? 
            normalizedRow['mãhànghóa'] ??
            normalizedRow['mãhh'] ??
            normalizedRow['mãsp'] ?? 
            normalizedRow['mãsảnphẩm'] ?? 
            normalizedRow['mãvậttư'] ??
            normalizedRow['itemno'] ??
            normalizedRow['itemcode'] ??
            normalizedRow['mã'] ??
            normalizedRow['item'] ??
            row.sku ?? row.SKU ?? row['Mã Hàng'] ?? ''
          ).trim();

          if (!sku) return;

          let name = String(
            normalizedRow['name'] ?? 
            normalizedRow['tênhàng'] ?? 
            normalizedRow['tênhànghóa'] ??
            normalizedRow['tênhh'] ??
            normalizedRow['tênsp'] ?? 
            normalizedRow['tênsảnphẩm'] ?? 
            normalizedRow['tênvậttư'] ??
            normalizedRow['productname'] ?? 
            normalizedRow['itemname'] ??
            normalizedRow['description'] ??
            normalizedRow['môtả'] ??
            normalizedRow['tên'] ??
            row.name ?? row.Name ?? row['Tên hàng'] ?? ''
          ).trim();

          const designationCode = normalizedRow['designationcode'] !== undefined ? String(normalizedRow['designationcode']).trim() : 
                                (normalizedRow['mãchỉđịnh'] !== undefined ? String(normalizedRow['mãchỉđịnh']).trim() : undefined);

          const loaiChiDinh = normalizedRow['loaichidinh'] !== undefined ? String(normalizedRow['loaichidinh']).trim() : 
                             (normalizedRow['loạichỉđịnh'] !== undefined ? String(normalizedRow['loạichỉđịnh']).trim() : undefined);

          const ghiChu = normalizedRow['ghichu'] !== undefined ? String(normalizedRow['ghichu']).trim() : 
                        (normalizedRow['ghichú'] !== undefined ? String(normalizedRow['ghichú']).trim() : undefined);

          const lotNo = normalizedRow['lotno'] !== undefined ? String(normalizedRow['lotno']).trim() : 
                       (normalizedRow['lotno.'] !== undefined ? String(normalizedRow['lotno.']).trim() : undefined);

          const existingIndex = skuToProductIndex.get(sku);
          const alreadyInNewIndex = skuToNewProductIndex.get(sku);
          
          if (existingIndex !== undefined) {
            let hasChanged = false;
            const existing = updatedProducts[existingIndex];
            if (name !== undefined && existing.name !== name && name !== '') {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], name };
              hasChanged = true;
            }
            if (designationCode !== undefined && existing.designationCode !== designationCode) {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], designationCode };
              hasChanged = true;
            }
            if (loaiChiDinh !== undefined && existing.loaiChiDinh !== loaiChiDinh) {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], loaiChiDinh };
              hasChanged = true;
            }
            if (ghiChu !== undefined && existing.ghiChu !== ghiChu) {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], ghiChu };
              hasChanged = true;
            }
            if (hasChanged) {
              changedProducts.push(updatedProducts[existingIndex]);
            }
          } else if (alreadyInNewIndex !== undefined) {
            if (name !== undefined && newProductsToAdd[alreadyInNewIndex].name === 'Sản phẩm mới' && name !== '') {
              newProductsToAdd[alreadyInNewIndex] = { ...newProductsToAdd[alreadyInNewIndex], name };
            }
            if (designationCode !== undefined && newProductsToAdd[alreadyInNewIndex].designationCode !== designationCode) {
              newProductsToAdd[alreadyInNewIndex] = { ...newProductsToAdd[alreadyInNewIndex], designationCode };
            }
            if (loaiChiDinh !== undefined && newProductsToAdd[alreadyInNewIndex].loaiChiDinh !== loaiChiDinh) {
              newProductsToAdd[alreadyInNewIndex] = { ...newProductsToAdd[alreadyInNewIndex], loaiChiDinh };
            }
            if (ghiChu !== undefined && newProductsToAdd[alreadyInNewIndex].ghiChu !== ghiChu) {
              newProductsToAdd[alreadyInNewIndex] = { ...newProductsToAdd[alreadyInNewIndex], ghiChu };
            }
          } else {
            const newProductIndex = newProductsToAdd.length;
            skuToNewProductIndex.set(sku, newProductIndex);
            newProductsToAdd.push({
              id: generateId(),
              sku: sku,
              name: name || 'Sản phẩm mới',
              category: normalizedRow['category'] || normalizedRow['loại'] || 'Tự động tạo',
              unit: normalizedRow['unit'] || normalizedRow['đơnvị'] || 'Cái',
              minStock: 0,
              lotNo: lotNo || '',
              ghiChu: ghiChu || '',
              designationCode: designationCode || '',
              loaiChiDinh: loaiChiDinh || ''
            });
          }
        });

        const finalProducts = [...updatedProducts, ...newProductsToAdd];
        const finalProductsMap = new Map(finalProducts.map(p => [p.sku, p]));

        let currentSkuForTransaction = '';

        setTransactions(prevTransactions => {
          const existingTransactionKeys = new Set(prevTransactions.map(t => 
            `${t.productId}|${t.type}|${t.quantity}|${t.date}|${t.partner}|${t.lotNo}|${t.designationCode}`
          ));

          const newTransactions: Transaction[] = data.map((row) => {
            const normalizedRow: any = {};
            Object.keys(row).forEach(key => {
              normalizedRow[normalizeKey(key)] = row[key];
            });

            let sku = String(
              normalizedRow['sku'] || 
              normalizedRow['mãhàng'] || 
              normalizedRow['mãhànghóa'] ||
              normalizedRow['mãhh'] ||
              normalizedRow['mãsp'] || 
              normalizedRow['mãsảnphẩm'] || 
              normalizedRow['mãvậttư'] ||
              normalizedRow['itemno'] ||
              normalizedRow['itemcode'] ||
              normalizedRow['mã'] ||
              normalizedRow['item'] ||
              row.sku || row.SKU || row['Mã Hàng'] || ''
            ).trim();

            if (!sku) {
              if (currentSkuForTransaction) {
                sku = currentSkuForTransaction;
              } else {
                return null;
              }
            } else {
              currentSkuForTransaction = sku;
            }

            const product = finalProductsMap.get(sku);
            if (!product) return null;

            const rawDate = normalizedRow['date'] || normalizedRow['ngày'] || normalizedRow['ngàynhập'] || normalizedRow['ngàyxuất'] || row.date || row.Date || row['Ngày nhập'] || row['Ngày xuất'];
            let formattedDate = format(new Date(), 'dd/MM/yyyy');
            
            if (rawDate) {
              if (typeof rawDate === 'number') {
                const date = XLSX.SSF.parse_date_code(rawDate);
                formattedDate = format(new Date(date.y, date.m - 1, date.d), 'dd/MM/yyyy');
              } else {
                formattedDate = String(rawDate).trim();
              }
            }

            const rowLoai = normalizedRow['loaichidinh'] !== undefined ? String(normalizedRow['loaichidinh']).trim() : 
                           (normalizedRow['loạichỉđịnh'] !== undefined ? String(normalizedRow['loạichỉđịnh']).trim() : undefined);
            const rowLot = normalizedRow['lotno'] !== undefined ? String(normalizedRow['lotno']).trim() : 
                          (normalizedRow['lotno.'] !== undefined ? String(normalizedRow['lotno.']).trim() : undefined);
            const rowGhiChu = normalizedRow['ghichu'] !== undefined ? String(normalizedRow['ghichu']).trim() : 
                             (normalizedRow['ghichú'] !== undefined ? String(normalizedRow['ghichú']).trim() : undefined);
            const rowDesignation = normalizedRow['designationcode'] !== undefined ? String(normalizedRow['designationcode']).replace(/\s+/g, '') : 
                                  (normalizedRow['mãchỉđịnh'] !== undefined ? String(normalizedRow['mãchỉđịnh']).replace(/\s+/g, '') : undefined);

            const quantity = Number(normalizedRow['quantity'] || normalizedRow['sốlượng'] || normalizedRow['sốlượngnhập'] || normalizedRow['sốlượngxuất'] || row.quantity || row.Quantity || row['Số lượng nhập'] || row['Số lượng xuất'] || 0);
            const partner = normalizedRow['partner'] || normalizedRow['đốitác'] || normalizedRow['kháchhàng'] || normalizedRow['nhàcungcấp'] || row.partner || row.Partner || 'N/A';
            const loaiChiDinh = rowLoai ?? '';
            const lotNo = rowLot ?? product?.lotNo ?? '';
            const ghiChu = rowGhiChu ?? product?.ghiChu ?? '';
            const designationCode = rowDesignation ?? product?.designationCode ?? '';

            const transactionKey = `${product.id}|${activeTab}|${quantity}|${formattedDate}|${partner}|${lotNo}|${designationCode}`;
            
            if (existingTransactionKeys.has(transactionKey)) {
              return null;
            }

            return {
              id: generateId(),
              productId: product.id,
              type: activeTab as 'inbound' | 'outbound',
              quantity,
              date: formattedDate,
              partner,
              loaiChiDinh,
              lotNo,
              ghiChu,
              designationCode
            };
          }).filter(Boolean) as Transaction[];

          if (newTransactions.length > 0) {
            api.transactions.upsertAll(newTransactions).catch(err => console.error('Error syncing transactions:', err));
            return [...prevTransactions, ...newTransactions];
          }
          return prevTransactions;
        });

        const productsToUpsert = [...changedProducts, ...newProductsToAdd];
        if (productsToUpsert.length > 0) {
          api.products.upsertAll(productsToUpsert).catch(err => console.error('Error syncing products:', err));
        }
        return finalProducts;
      });
    } else if (activeTab === 'customers') {
      setCustomers(prev => {
        const newCustomers: Customer[] = [];
        const currentCustomers = [...prev];
        
        data.forEach((row, index) => {
          const normalizedRow: any = {};
          Object.keys(row).forEach(key => {
            normalizedRow[key.toLowerCase().trim()] = row[key];
          });

          const code = String(
            normalizedRow['code'] || 
            normalizedRow['mã'] || 
            normalizedRow['mã khách hàng'] || 
            normalizedRow['mã kh'] || 
            normalizedRow['customer code'] || 
            normalizedRow['cust code'] ||
            row['Code'] || row['code'] || row['Mã'] || ''
          ).trim();
          
          const name = String(
            normalizedRow['name'] || 
            normalizedRow['tên'] || 
            normalizedRow['tên khách hàng'] || 
            normalizedRow['tên kh'] || 
            normalizedRow['customer name'] || 
            normalizedRow['cust name'] ||
            row['Name'] || row['name'] || row['Tên'] || ''
          ).trim();
          
          if (code && name) {
            const existsInCurrent = currentCustomers.find(c => 
              c.code.toLowerCase() === code.toLowerCase() && 
              c.name.toLowerCase() === name.toLowerCase()
            );
            const existsInNew = newCustomers.find(c => 
              c.code.toLowerCase() === code.toLowerCase() && 
              c.name.toLowerCase() === name.toLowerCase()
            );
            
            if (!existsInCurrent && !existsInNew) {
              newCustomers.push({
                id: generateId(),
                code,
                name
              });
            }
          }
        });

        if (newCustomers.length === 0) return prev;
        const updated = [...prev, ...newCustomers];
        api.customers.upsertAll(newCustomers).catch(err => console.error('Error syncing customers:', err));
        return updated;
      });
    }
  };

  const parseQRCode = (qrcode: string) => {
    if (!qrcode) return null;
    
    // Clean trailing hyphen from the whole string if it exists
    let cleanQRCode = qrcode.trim();
    if (cleanQRCode.endsWith('-')) {
      cleanQRCode = cleanQRCode.slice(0, -1);
    }

    const parts = cleanQRCode.split('|');
    
    if (parts.length < 2) {
      let sku = cleanQRCode;
      if (sku.endsWith('-')) sku = sku.slice(0, -1);
      return { sku, partner: '', date: '' };
    }
    
    let sku = parts[0].trim();
    if (sku.endsWith('-')) {
      sku = sku.slice(0, -1);
    }
    
    const rest = parts[1].trim();
    const restParts = rest.split('-');
    
    const partner = restParts[0] || '';
    let date = restParts[1] || '';
    
    if (date && date.includes('/')) {
      const dateParts = date.split('/');
      if (dateParts[2] && dateParts[2].length === 2) {
        dateParts[2] = '20' + dateParts[2];
        date = dateParts.join('/');
      }
    }
    
    return { sku, partner, date };
  };

  const processLocationStockSheet = (sheetData: any[][]) => {
    let currentSection: 'INPUT' | 'OUTPUT' | null = null;
    const inputEntries: any[] = [];
    const outputQRCodes: Set<string> = new Set();

    sheetData.forEach(row => {
      if (!row || row.length === 0) return;
      const firstCell = String(row[0] || '').trim().toUpperCase();
      
      if (firstCell === 'INPUT') {
        currentSection = 'INPUT';
        return;
      }
      if (firstCell === 'OUTPUT') {
        currentSection = 'OUTPUT';
        return;
      }

      if (currentSection === 'INPUT') {
        const qrcode = String(row[0] || '').trim();
        if (qrcode && qrcode !== 'QRCODE') {
          const parsed = parseQRCode(qrcode);
          inputEntries.push({
            qrcode,
            sku: String(row[1] || '').trim() || parsed?.sku || '',
            partner: String(row[2] || '').trim() || parsed?.partner || '',
            date: String(row[3] || '').trim() || parsed?.date || '',
            note: String(row[4] || '').trim(),
            location: String(row[5] || '').trim(),
            quantity: parseInt(String(row[6] || '1')) || 1
          });
        }
      } else if (currentSection === 'OUTPUT') {
        const qrcode = String(row[0] || '').trim();
        if (qrcode && qrcode !== 'QRCODE') {
          outputQRCodes.add(qrcode);
        }
      }
    });

    const filtered = inputEntries.filter(entry => !outputQRCodes.has(entry.qrcode));
    
    const grouped = new Map<string, any>();
    filtered.forEach(entry => {
      const key = `${entry.qrcode}|${entry.location}`;
      if (!grouped.has(key)) {
        grouped.set(key, { ...entry });
      } else {
        const existing = grouped.get(key);
        if (!existing.sku) existing.sku = entry.sku;
        if (!existing.partner) existing.partner = entry.partner;
        if (!existing.date) existing.date = entry.date;
        if (!existing.note) existing.note = entry.note;
        existing.quantity = (existing.quantity || 0) + (entry.quantity || 0);
      }
    });

    return {
      entries: Array.from(grouped.values()),
      outputQRCodes: Array.from(outputQRCodes)
    };
  };

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const fileName = file.name.toLowerCase();
    
    if (fileName.endsWith('.csv')) {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          if (activeTab === 'deliveryNote') {
            processDeliveryNoteData(results.data);
          } else if (activeTab === 'customers') {
            processCustomerData(results.data);
          } else if (activeTab === 'location') {
            processLocationData(results.data);
          } else {
            processData(results.data);
          }
          if (fileInputRef.current) fileInputRef.current.value = '';
        }
      });
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.xlsm')) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        if (activeTab === 'location' && locationSubTab === 'inventory') {
          const sheetName = workbook.SheetNames[2] || workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
          const { entries: processedData, outputQRCodes } = processLocationStockSheet(sheetData);
          
          const newEntries: LocationEntry[] = processedData.map((item, index) => ({
            id: generateId(),
            ...item,
            type: 'inventory'
          }));
          
          setLocationInventoryEntries(prev => {
            const grouped = new Map<string, LocationEntry>();
            prev.forEach(entry => {
              const key = `${entry.qrcode}|${entry.location || ''}`;
              if (!outputQRCodes.includes(entry.qrcode)) {
                grouped.set(key, { ...entry });
              }
            });
            
            newEntries.forEach(entry => {
              const key = `${entry.qrcode}|${entry.location || ''}`;
              if (grouped.has(key)) {
                const existing = grouped.get(key)!;
                if (!existing.sku) existing.sku = entry.sku;
                if (!existing.partner) existing.partner = entry.partner;
                if (!existing.date) existing.date = entry.date;
                if (!existing.note) existing.note = entry.note;
                existing.quantity = (existing.quantity || 0) + (entry.quantity || 0);
              } else {
                grouped.set(key, entry);
              }
            });
            
            const finalEntries = Array.from(grouped.values());
            
            (async () => {
              try {
                const entriesToDelete = prev.filter(e => outputQRCodes.includes(e.qrcode));
                if (entriesToDelete.length > 0) {
                  await api.locationEntries.deleteMany(entriesToDelete.map(e => e.id));
                }
                
                // Only upsert changed items
                const changedEntries = finalEntries.filter(e => {
                  const existing = prev.find(p => p.id === e.id);
                  if (!existing) return true;
                  return JSON.stringify(existing) !== JSON.stringify(e);
                });

                if (changedEntries.length > 0) {
                  await api.locationEntries.upsertAll(changedEntries);
                }
              } catch (error) {
                console.error('Error syncing Excel location data:', error);
              }
            })();

            return finalEntries;
          });
        } else {
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
          
          if (activeTab === 'deliveryNote') {
            processDeliveryNoteData(jsonData);
          } else if (activeTab === 'customers') {
            processCustomerData(jsonData);
          } else if (activeTab === 'location') {
            processLocationData(jsonData);
          } else {
            processData(jsonData);
          }
        }
        
        if (fileInputRef.current) fileInputRef.current.value = '';
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const processDeliveryNoteData = (data: any[]) => {
    const normalizeKey = (key: string) => {
      return key.toLowerCase().trim()
        .replace(/\s+/g, '');
    };

    const mappedData: DeliveryNoteItem[] = data.map((row, index) => {
      const normalizedRow: any = {};
      Object.keys(row).forEach(key => {
        normalizedRow[normalizeKey(key)] = row[key];
      });

      const itemNo = String(
        normalizedRow['itemno'] || 
        normalizedRow['item'] || 
        normalizedRow['mãhàng'] ||
        row['Item No.'] || row['item'] || row['ITEM'] || ''
      ).trim();

      const ovnSaleOrder = String(
        normalizedRow['ovnsaleorder'] || 
        row['OVN Sale Order'] || ''
      ).trim();

      const ovnProductionOrder = String(
        normalizedRow['ovnproductionorder'] || 
        row['OVN Production Order'] || ''
      ).trim();

      const qtyErp = Number(
        normalizedRow['quantity'] || 
        normalizedRow['qtyerp'] || 
        normalizedRow['sốlượng'] ||
        row['Quantity'] || row['Qty ERP'] || 0
      );
      
      const noValue = String(
        normalizedRow['no.'] || 
        normalizedRow['no'] || 
        row['No.'] || row['NO.'] || row['No'] || row['NO'] || ''
      ).trim();

      const customerCode = String(
        normalizedRow['customercode'] || 
        normalizedRow['mãkháchhàng'] || 
        normalizedRow['selltocustomername'] ||
        row['Customer code'] || row['Sell-to Customer Name'] || ''
      ).trim();

      let finalNoCode = noValue;
      if (!finalNoCode && customerCode) {
        const foundCustomer = customers.find(c => 
          c.name.toLowerCase() === customerCode.toLowerCase() || 
          c.code.toLowerCase() === customerCode.toLowerCase()
        );
        if (foundCustomer) {
          finalNoCode = foundCustomer.code;
        }
      }

      const loaiChiDinhFromRow = String(
        normalizedRow['loaichidinh'] || 
        normalizedRow['loạichỉđịnh'] || 
        row['Loại chỉ định'] || ''
      ).trim().toUpperCase();

      const finalLoaiChiDinh = determineLoaiChiDinh(itemNo, ovnProductionOrder, ovnSaleOrder, finalNoCode, loaiChiDinhFromRow);

      const materialName = String(
        normalizedRow['materialname'] || 
        normalizedRow['tênhàng'] ||
        normalizedRow['ovnfullname'] ||
        row['Material Name'] || row['OVN Full Name'] || ''
      ).trim();

      const unit = String(
        normalizedRow['unit'] || 
        normalizedRow['đơnvị'] ||
        row['Unit'] || 'YDS'
      ).trim();

      const brand = String(
        normalizedRow['brand'] || 
        normalizedRow['brandcode'] ||
        row['Brand'] || row['Brand Code'] || ''
      ).trim();

      const finalDestination = String(
        normalizedRow['finaldestination'] || 
        normalizedRow['đến'] || 
        row['Final Destination'] || ''
      ).trim();

      const remark = String(
        normalizedRow['remark'] || 
        normalizedRow['ghichú'] || 
        row['remark'] || ''
      ).trim();

      const lotNo = String(
        normalizedRow['lotno'] || 
        normalizedRow['lot'] || 
        row['Lot No'] || row['lot'] || ''
      ).trim();

      return {
        id: generateId(),
        no: index + 1,
        ovnSaleOrder,
        ovnProductionOrder,
        item: itemNo,
        materialName,
        unit,
        qtyErp,
        actualQty: 0,
        lotNo: lotNo,
        actualIssuedQty: 0,
        remark,
        loaiChiDinh: finalLoaiChiDinh,
        brand,
        customerCode,
        finalDestination,
        noCode: finalNoCode,
        location: getLocationByItemAndLot(itemNo, lotNo),
        stock: ''
      };
    });

    const finalData = recalculateActualQty(mappedData.map((item, index) => ({
      ...item,
      no: index + 1
    })).sort((a, b) => {
      if (a.item !== b.item) return a.item.localeCompare(b.item);
      return (a.loaiChiDinh || 'NORMAL').localeCompare(b.loaiChiDinh || 'NORMAL');
    }));

    setDeliveryNotes(finalData);
    api.deliveryNotes.upsertAll(finalData).catch(err => console.error('Error syncing delivery notes:', err));
    showNotification(`Đã tải lên ${mappedData.length} dòng. Hãy nhấn "Gắn Lot tự động" để tìm tồn kho phù hợp.`, 'success');
  };

  const processCustomerData = (data: any[]) => {
    setCustomers(prev => {
      const newCustomers: Customer[] = [];
      const currentCustomers = [...prev];
      
      data.forEach((row, index) => {
        const normalizedRow: any = {};
        Object.keys(row).forEach(key => {
          normalizedRow[key.toLowerCase().trim()] = row[key];
        });

        const code = String(
          normalizedRow['code'] ?? 
          normalizedRow['mã'] ?? 
          normalizedRow['mã khách hàng'] ?? 
          normalizedRow['mã kh'] ?? 
          normalizedRow['customer code'] ?? 
          normalizedRow['cust code'] ??
          row['Code'] ?? row['code'] ?? row['Mã'] ?? ''
        ).trim();
        
        const name = String(
          normalizedRow['name'] ?? 
          normalizedRow['tên'] ?? 
          normalizedRow['tên khách hàng'] ?? 
          normalizedRow['tên kh'] ?? 
          normalizedRow['customer name'] ?? 
          normalizedRow['cust name'] ??
          row['Name'] ?? row['name'] ?? row['Tên'] ?? ''
        ).trim();
        
        if (code && name) {
          const existsInCurrent = currentCustomers.find(c => 
            c.code.toLowerCase() === code.toLowerCase() && 
            c.name.toLowerCase() === name.toLowerCase()
          );
          const existsInNew = newCustomers.find(c => 
            c.code.toLowerCase() === code.toLowerCase() && 
            c.name.toLowerCase() === name.toLowerCase()
          );
          
          if (!existsInCurrent && !existsInNew) {
            newCustomers.push({
              id: generateId(),
              code,
              name
            });
          }
        }
      });

      if (newCustomers.length === 0) return prev;
      const updated = [...prev, ...newCustomers];
      Promise.all(newCustomers.map(c => api.customers.upsert(c))).catch(err => console.error('Error syncing customers:', err));
      return updated;
    });
  };

  const chartData = useMemo(() => {
    const dataMap: Record<string, number> = {};
    inventory.forEach(item => {
      const type = item.loaiChiDinh?.toUpperCase() || 'NORMAL';
      let category = 'NORMAL';
      if (type.includes('KH')) category = 'KH';
      else if (type.includes('SO')) category = 'SO';
      else if (type.includes('NCC')) category = 'NCC';
      else if (type.includes('VAI')) category = 'VAI KHCC';
      
      dataMap[category] = (dataMap[category] || 0) + item.currentStock;
    });
    return Object.entries(dataMap).map(([name, value]) => ({ name, value }));
  }, [inventory]);

  const stats = {
    totalProducts: products.length,
    lowStock: inventory.filter(item => item.currentStock <= item.minStock).length,
    totalInboundQty: inventory.reduce((sum, item) => sum + item.totalInbound, 0),
    totalOutboundQty: inventory.reduce((sum, item) => sum + item.totalOutbound, 0),
    chiDinhKH: inventory.filter(item => item.loaiChiDinh?.toUpperCase().includes('KH')).length,
    chiDinhSO: inventory.filter(item => item.loaiChiDinh?.toUpperCase().includes('SO')).length,
    chiDinhNCC: inventory.filter(item => item.loaiChiDinh?.toUpperCase().includes('NCC')).length,
    normal: inventory.filter(item => !item.loaiChiDinh || item.loaiChiDinh?.toUpperCase().includes('NORMAL') || item.loaiChiDinh?.toUpperCase().includes('VAI')).length,
    totalStockQty: inventory.reduce((sum, item) => sum + item.currentStock, 0),
    totalStockCodes: inventory.filter(item => item.currentStock > 0).length,
  };

  return (
    <div className="min-h-screen bg-[#E4E3E0] text-[#141414] font-sans flex relative">
      {/* Mobile Sidebar Overlay */}
      <AnimatePresence>
        {isSidebarOpen && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={() => setIsSidebarOpen(false)}
            className="fixed inset-0 bg-black/50 z-40 lg:hidden"
          />
        )}
      </AnimatePresence>

      {/* Sidebar */}
      <aside className={cn(
        "fixed inset-y-0 left-0 z-50 w-64 border-r border-[#141414] bg-[#E4E3E0] flex flex-col transition-transform duration-300 lg:relative lg:translate-x-0 no-print",
        isSidebarOpen ? "translate-x-0" : "-translate-x-full"
      )}>
        <div className="p-6 border-b border-[#141414] flex justify-between items-center">
          <div>
            <h1 className="font-serif italic text-2xl font-bold tracking-tight">KHO.LOG</h1>
            <p className="text-[10px] uppercase tracking-widest opacity-50 mt-1">Warehouse Management System</p>
          </div>
          <button onClick={() => setIsSidebarOpen(false)} className="lg:hidden p-2">
            <X size={20} />
          </button>
        </div>

        <nav className="flex-1 px-4 py-8 space-y-2 overflow-y-auto">
          <NavItem 
            active={activeTab === 'dashboard'} 
            onClick={() => { setActiveTab('dashboard'); setIsSidebarOpen(false); }}
            icon={<LayoutDashboard size={18} />}
            label="Tổng quan"
          />
          <NavItem 
            active={activeTab === 'inbound'} 
            onClick={() => { setActiveTab('inbound'); setIsSidebarOpen(false); }}
            icon={<ArrowDownToLine size={18} />}
            label="Nhập kho"
          />
          <NavItem 
            active={activeTab === 'outbound'} 
            onClick={() => { setActiveTab('outbound'); setIsSidebarOpen(false); }}
            icon={<ArrowUpFromLine size={18} />}
            label="Xuất kho"
          />
          <NavItem 
            active={activeTab === 'inventory'} 
            onClick={() => { setActiveTab('inventory'); setIsSidebarOpen(false); }}
            icon={<Package size={18} />}
            label="Tồn kho"
          />
          <NavItem 
            active={activeTab === 'customers'} 
            onClick={() => { setActiveTab('customers'); setIsSidebarOpen(false); }}
            icon={<Users size={18} />}
            label="Danh sách khách hàng"
          />
          <NavItem 
            active={activeTab === 'deliveryNote'} 
            onClick={() => { setActiveTab('deliveryNote'); setIsSidebarOpen(false); }}
            icon={<FileText size={18} />}
            label="Lệnh xuất kho"
          />
          <NavItem 
            active={activeTab === 'location'} 
            onClick={() => { setActiveTab('location'); setIsSidebarOpen(false); }}
            icon={<MapPin size={18} />}
            label="Vị Trí"
          />
        </nav>

        <div className="px-6 py-4 border-t border-[#141414]/10">
          <StorageUsageBar showNotification={showNotification} />
        </div>

        <div className="p-6 border-t border-[#141414]">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-full bg-[#141414] text-[#E4E3E0] flex items-center justify-center text-xs font-bold">
              AD
            </div>
            <div>
              <p className="text-xs font-bold">Admin User</p>
              <p className="text-[10px] opacity-50">quanlykhotheovitri@gmail.com</p>
            </div>
          </div>
        </div>
      </aside>

      <main className="flex-1 flex flex-col overflow-hidden w-full">
        <header className="h-20 border-b border-[#141414] flex items-center justify-between px-4 sm:px-8 bg-[#E4E3E0]/80 backdrop-blur-sm z-10 gap-4 no-print">
          <div className="flex items-center gap-4 flex-1">
            <button 
              onClick={() => setIsSidebarOpen(true)}
              className="lg:hidden p-2 border border-[#141414] hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
            >
              <Menu size={20} />
            </button>
            <div className="flex items-center gap-3 flex-1 max-w-md">
              <Search size={18} className="opacity-40 shrink-0" />
              <input 
                type="text" 
                placeholder="Tìm kiếm..." 
                className="bg-transparent border-none outline-none w-full text-sm placeholder:italic"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
              />
            </div>
          </div>

          <div className="flex items-center gap-2 sm:gap-4">
            <div className="hidden sm:flex items-center gap-2 sm:gap-4">
              <div 
                className={cn(
                  "flex items-center gap-2 px-3 sm:px-4 py-2 border transition-colors cursor-default",
                  !isOnline 
                    ? "border-red-500 text-red-600 bg-red-50" 
                    : isSaving 
                      ? "border-amber-500 text-amber-600 bg-amber-50" 
                      : "border-green-500 text-green-600 bg-green-50"
                )}
                title={!isOnline ? "Hoạt động Ngoại Tuyến (Offline)" : isSaving ? "Đang đồng bộ..." : "Đã đồng bộ với Supabase"}
              >
                {!isOnline ? (
                  <WifiOff size={14} />
                ) : isSaving ? (
                  <div className="w-3 h-3 border-2 border-amber-600 border-t-transparent rounded-full animate-spin" />
                ) : (
                  <CheckCircle2 size={14} />
                )}
                <span className="hidden md:inline text-[10px] font-bold uppercase tracking-wider">
                  {!isOnline ? 'Ngoại tuyến' : isSaving ? 'Đang lưu...' : 'Đã lưu'}
                </span>
              </div>
              
              {selectedRows.length > 0 && (
                <button 
                  onClick={() => {
                    const type = activeTab === 'inventory' ? 'product' : 
                                 activeTab === 'customers' ? 'customer' : 
                                 activeTab === 'location' ? 'location' : 
                                 'transaction';
                    setDeleteTarget({ id: 'bulk', type });
                    setIsDeleteConfirmOpen(true);
                  }}
                  className="flex items-center gap-2 px-3 sm:px-4 py-2 bg-red-600 text-white text-[10px] font-bold uppercase tracking-wider hover:bg-red-700 transition-colors"
                >
                  <Trash2 size={14} />
                  <span className="hidden md:inline">Xóa</span> ({selectedRows.length})
                </button>
              )}
            </div>

            <div className="flex items-center gap-2">
              <button 
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center gap-2 px-3 sm:px-4 py-2 border border-[#141414] text-[10px] font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                title="Cập nhật file"
              >
                <FileUp size={14} />
                <span className="inline">Cập nhật file</span>
              </button>
              <button className="p-2 border border-[#141414] hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors">
                <Filter size={18} />
              </button>
            </div>
          </div>
        </header>

        {!isSupabaseConfigured && (
          <div className="bg-amber-100 border-b border-amber-200 px-8 py-2 flex items-center justify-between">
            <div className="flex items-center gap-3 text-amber-800 text-[10px]">
              <AlertTriangle size={14} />
              <span>
                <strong>Supabase chưa được cấu hình:</strong> Vui lòng thiết lập <code>VITE_SUPABASE_URL</code> và <code>VITE_SUPABASE_ANON_KEY</code> trong menu Settings để đồng bộ dữ liệu.
              </span>
            </div>
            <button 
              onClick={() => setIsSetupModalOpen(true)}
              className="text-[10px] font-bold uppercase text-amber-900 hover:underline"
            >
              Hướng dẫn thiết lập
            </button>
          </div>
        )}

        <div className="flex-1 overflow-y-auto p-4 sm:p-8">
          <AnimatePresence mode="wait">
            {activeTab === 'dashboard' && (
              <motion.div 
                key="dashboard"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-8"
              >
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 sm:gap-6">
                  <StatCard 
                    label="CHI ĐỊNH THEO KH" 
                    value={stats.chiDinhKH} 
                    bgColor="bg-[#1a5f7a]" 
                    textColor="text-white"
                  />
                  <StatCard 
                    label="CHI ĐỊNH THEO SO" 
                    value={stats.chiDinhSO} 
                    bgColor="bg-[#fde5d4]" 
                    textColor="text-red-600"
                  />
                  <StatCard 
                    label="CHI ĐỊNH THEO NCC" 
                    value={stats.chiDinhNCC} 
                    bgColor="bg-[#f5d5f5]" 
                    textColor="text-red-600"
                  />
                  <StatCard 
                    label="NORMAL & VAI KHCC" 
                    value={stats.normal} 
                    bgColor="bg-[#d9e9f9]" 
                    textColor="text-red-600"
                  />
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
                  <div className="lg:col-span-3 border border-[#141414] p-4 sm:p-6 bg-[#1a5f7a] text-white">
                    <div className="border border-white/30 p-2 inline-block mb-6">
                      <h3 className="text-[10px] sm:text-xs font-bold uppercase tracking-widest">BIỂU ĐỒ TỒN KHO THEO LOẠI CHỈ ĐỊNH</h3>
                    </div>
                    <div className="h-[300px] sm:h-[400px]">
                      <ResponsiveContainer width="100%" height="100%" minWidth={0} minHeight={0}>
                        <BarChart data={chartData}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#ffffff" strokeOpacity={0.1} />
                          <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#ffffff' }} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#ffffff' }} />
                          <Tooltip 
                            contentStyle={{ backgroundColor: '#ffffff', color: '#141414', border: 'none' }}
                            itemStyle={{ color: '#141414' }}
                          />
                          <Bar dataKey="value" fill="#ffffff" radius={[4, 4, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div className="border border-[#141414] p-6 sm:p-8 bg-[#1a5f7a] text-white flex flex-col gap-4">
                    <div className="space-y-1">
                      <p className="text-[10px] sm:text-[11px] font-bold uppercase tracking-wider opacity-80">TỔNG MÃ TỒN KHO :</p>
                      <p className="text-2xl sm:text-3xl font-bold">{stats.totalStockCodes.toLocaleString()}</p>
                    </div>
                    <div className="space-y-1 mt-4">
                      <p className="text-[11px] font-bold uppercase tracking-wider opacity-80">TỔNG SỐ LƯỢNG TỒN KHO:</p>
                      <p className="text-3xl font-bold">{stats.totalStockQty.toLocaleString()}</p>
                    </div>
                  </div>
                </div>
              </motion.div>
            )}

            {(activeTab === 'inbound' || activeTab === 'outbound') && (
              <motion.div 
                key={activeTab}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-6"
              >
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <h2 className="font-serif italic text-2xl capitalize">
                    {activeTab === 'inbound' ? 'Danh sách nhập kho' : 'Danh sách xuất kho'}
                  </h2>
                  <div className="flex flex-wrap gap-2 sm:gap-3">
                    <div className="relative flex-1 sm:flex-none">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                      <input
                        type="text"
                        placeholder="Tìm kiếm..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="pl-10 pr-4 py-2 border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-full sm:w-64"
                      />
                    </div>
                    {selectedRows.length > 0 && (
                      <button
                        onClick={() => {
                          setDeleteTarget({ id: 'bulk', type: 'transaction' });
                          setIsDeleteConfirmOpen(true);
                        }}
                        className="flex items-center gap-2 px-4 py-2 bg-red-600 text-white text-xs font-bold uppercase tracking-wider hover:bg-red-700 transition-colors"
                      >
                        <Trash2 size={14} />
                        <span className="hidden sm:inline">Xóa</span> ({selectedRows.length})
                      </button>
                    )}
                    <button 
                      onClick={handleExportTransactions}
                      className="flex items-center gap-2 px-4 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                    >
                      <Download size={14} />
                      <span className="hidden sm:inline">Xuất Excel</span>
                    </button>
                    <button 
                      onClick={() => {
                        setEditingId(null);
                        setNewTransaction({ productId: '', type: activeTab as 'inbound' | 'outbound', quantity: 0, date: format(new Date(), 'dd/MM/yyyy'), partner: '', loaiChiDinh: '', lotNo: '', ghiChu: '', designationCode: '' });
                        setIsTransactionModalOpen(true);
                      }}
                      className="flex items-center gap-2 px-4 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <Plus size={14} />
                      <span className="hidden sm:inline">Thêm giao dịch</span>
                    </button>
                  </div>
                </div>

                <div className="border border-[#141414] overflow-x-auto max-h-[70vh]">
                  <table className="w-full border-collapse min-w-[1000px]">
                    <thead className="sticky top-0 z-10 shadow-sm">
                      <tr className="bg-[#001F3F] text-white text-[11px] uppercase tracking-wider">
                        <th className="border border-[#141414] p-3 w-10">
                          <button onClick={() => {
                            const currentIds = filteredTransactions.filter(t => t.type === activeTab).map(t => t.id);
                            if (selectedRows.length === currentIds.length) setSelectedRows([]);
                            else setSelectedRows(currentIds);
                          }}>
                            {selectedRows.length > 0 ? <CheckSquare size={14} /> : <Square size={14} />}
                          </button>
                        </th>
                        <th className="border border-[#141414] p-3 text-left">Mã Hàng</th>
                        <th className="border border-[#141414] p-3 text-left">Tên hàng</th>
                        <th className="border border-[#141414] p-3 text-left">Lot no</th>
                        <th className="border border-[#141414] p-3 text-center bg-yellow-400 text-[#141414]">
                          {activeTab === 'inbound' ? 'Số lượng nhập' : 'Số lượng xuất'}
                        </th>
                        <th className="border border-[#141414] p-3 text-left">
                          {activeTab === 'inbound' ? 'Ngày nhập' : 'Ngày xuất'}
                        </th>
                        <th className="border border-[#141414] p-3 text-left">Loại chỉ định</th>
                        <th className="border border-[#141414] p-3 text-left">Ghi chú</th>
                        <th className="border border-[#141414] p-3 text-left">Mã chỉ định</th>
                        <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                          Thao tác
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredTransactions
                        .filter(t => t.type === activeTab)
                        .map(t => {
                          const product = products.find(p => p.id === t.productId);
                          const isSelected = selectedRows.includes(t.id);
                          return (
                            <tr key={t.id} className={cn(
                              "bg-white text-xs hover:bg-gray-50 transition-colors",
                              isSelected && "bg-blue-50"
                            )}>
                              <td className="border border-[#141414] p-3 text-center">
                                <button onClick={() => toggleRowSelection(t.id)}>
                                  {isSelected ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} />}
                                </button>
                              </td>
                              <td className="border border-[#141414] p-3 font-mono">{product?.sku}</td>
                              <td className="border border-[#141414] p-3 font-bold">{product?.name}</td>
                              <td className="border border-[#141414] p-3 italic">{t.lotNo ?? product?.lotNo ?? ''}</td>
                              <td className="border border-[#141414] p-3 text-center font-bold">{t.quantity}</td>
                              <td className="border border-[#141414] p-3">{t.date}</td>
                              <td className="border border-[#141414] p-3 opacity-60">{t.loaiChiDinh}</td>
                              <td className="border border-[#141414] p-3">{t.ghiChu ?? product?.ghiChu ?? ''}</td>
                              <td className="border border-[#141414] p-3 font-mono">{t.designationCode ?? product?.designationCode ?? ''}</td>
                              <td className="border border-[#141414] p-3 text-center">
                                <div className="flex items-center justify-center gap-2">
                                  <button 
                                    onClick={() => {
                                      setEditingId(t.id);
                                      setNewTransaction(t);
                                      setIsTransactionModalOpen(true);
                                    }}
                                    className="p-1 hover:bg-gray-200 rounded transition-colors"
                                  >
                                    <Edit2 size={14} className="text-blue-600" />
                                  </button>
                                  <button 
                                    onClick={() => {
                                      setDeleteTarget({ id: t.id, type: 'transaction' });
                                      setIsDeleteConfirmOpen(true);
                                    }}
                                    className="p-1 hover:bg-gray-200 rounded transition-colors"
                                  >
                                    <Trash2 size={14} className="text-red-600" />
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                  </table>
                </div>
              </motion.div>
            )}
            {activeTab === 'inventory' && (
              <motion.div 
                key="inventory"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-6"
              >
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <h2 className="font-serif italic text-2xl">Báo cáo tồn kho</h2>
                  <div className="flex flex-wrap gap-2 sm:gap-3">
                    <div className="relative flex-1 sm:flex-none">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                      <input
                        type="text"
                        placeholder="Tìm kiếm..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="pl-10 pr-4 py-2 border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-full sm:w-64"
                      />
                    </div>
                    {selectedRows.length > 0 && (
                      <button
                        onClick={() => {
                          setDeleteTarget({ id: 'bulk', type: 'inventory' });
                          setIsDeleteConfirmOpen(true);
                        }}
                        className="flex items-center gap-2 px-4 py-2 bg-red-600 text-white text-xs font-bold uppercase tracking-wider hover:bg-red-700 transition-colors"
                      >
                        <Trash2 size={14} />
                        <span className="hidden sm:inline">Xóa</span> ({selectedRows.length})
                      </button>
                    )}
                    <button onClick={handleExportInventory}  className="flex items-center gap-2 px-4 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors">
                      <Download size={14} />
                      <span className="hidden sm:inline">Xuất báo cáo</span>
                    </button>
                    <button 
                      onClick={() => {
                        setEditingId(null);
                        setNewProduct({ sku: '', name: '', category: '', unit: '', minStock: 0, lotNo: '', ghiChu: '', designationCode: '', loaiChiDinh: '', quantity: 0 });
                        setIsProductModalOpen(true);
                      }}
                      className="flex items-center gap-2 px-4 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <Plus size={14} />
                      <span className="hidden sm:inline">Thêm sản phẩm</span>
                    </button>
                  </div>
                </div>

                <div className="border border-[#141414] overflow-x-auto max-h-[70vh]">
                  <table className="w-full border-collapse min-w-[1200px]">
                    <thead className="sticky top-0 z-10 shadow-sm">
                      <tr className="bg-[#001F3F] text-white text-[11px] uppercase tracking-wider">
                        <th className="border border-[#141414] p-3 w-10">
                          <button onClick={() => {
                            const currentIds = filteredInventory.map(i => i.id);
                            if (selectedRows.length === currentIds.length) setSelectedRows([]);
                            else setSelectedRows(currentIds);
                          }}>
                            {selectedRows.length > 0 ? <CheckSquare size={14} /> : <Square size={14} />}
                          </button>
                        </th>
                        <th className="border border-[#141414] p-3 text-left">Mã Hàng</th>
                        <th className="border border-[#141414] p-3 text-left">Tên hàng</th>
                        <th className="border border-[#141414] p-3 text-left">Lot no</th>
                        <th className="border border-[#141414] p-3 text-center bg-yellow-400 text-[#141414]">Số lượng nhập</th>
                        <th className="border border-[#141414] p-3 text-center bg-yellow-400 text-[#141414]">Số lượng xuất</th>
                        <th className="border border-[#141414] p-3 text-center bg-yellow-400 text-[#141414]">Tồn cuối</th>
                        <th className="border border-[#141414] p-3 text-left">Loại chỉ định</th>
                        <th className="border border-[#141414] p-3 text-left">Ghi chú</th>
                        <th className="border border-[#141414] p-3 text-left">Mã chỉ định</th>
                        <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                          Thao tác
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredInventory.map(item => {
                        const isSelected = selectedRows.includes(item.id);
                        return (
                          <tr key={item.id} className={cn(
                            "bg-white text-xs hover:bg-gray-50 transition-colors",
                            isSelected && "bg-blue-50"
                          )}>
                            <td className="border border-[#141414] p-3 text-center">
                              <button onClick={() => toggleRowSelection(item.id)}>
                                {isSelected ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} />}
                              </button>
                            </td>
                            <td className="border border-[#141414] p-3 font-mono">{item.sku}</td>
                            <td className="border border-[#141414] p-3 font-bold">{item.name}</td>
                            <td className="border border-[#141414] p-3 italic">{item.lotNo}</td>
                            <td className="border border-[#141414] p-3 text-center font-bold">{item.totalInbound}</td>
                            <td className="border border-[#141414] p-3 text-center font-bold">{item.totalOutbound}</td>
                            <td className="border border-[#141414] p-3 text-center font-bold text-blue-600">{item.currentStock}</td>
                            <td className="border border-[#141414] p-3 opacity-60">{item.loaiChiDinh}</td>
                            <td className="border border-[#141414] p-3">{item.ghiChu}</td>
                            <td className="border border-[#141414] p-3 font-mono">{item.designationCode}</td>
                            <td className="border border-[#141414] p-3 text-center">
                              <div className="flex items-center justify-center gap-2">
                                <button 
                                  onClick={() => {
                                    setEditingId(item.productId);
                                    setNewProduct({
                                      ...item,
                                      quantity: item.currentStock
                                    });
                                    setIsProductModalOpen(true);
                                  }}
                                  className="p-1 hover:bg-gray-200 rounded transition-colors"
                                >
                                  <Edit2 size={14} className="text-blue-600" />
                                </button>
                                <button 
                                  onClick={() => {
                                    setDeleteTarget({ id: item.productId, type: 'product' });
                                    setIsDeleteConfirmOpen(true);
                                  }}
                                  className="p-1 hover:bg-gray-200 rounded transition-colors"
                                >
                                  <Trash2 size={14} className="text-red-600" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </motion.div>
            )}
            {activeTab === 'customers' && (
              <motion.div 
                key="customers"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-6"
              >
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <h2 className="font-serif italic text-2xl">Danh sách khách hàng</h2>
                  <div className="flex flex-wrap gap-2 sm:gap-3">
                    <div className="relative flex-1 sm:flex-none">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                      <input
                        type="text"
                        placeholder="Tìm kiếm..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="pl-10 pr-4 py-2 border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-full sm:w-64"
                      />
                    </div>
                    {selectedRows.length > 0 && (
                      <button
                        onClick={() => {
                          setDeleteTarget({ id: 'bulk', type: 'customer' });
                          setIsDeleteConfirmOpen(true);
                        }}
                        className="flex items-center gap-2 px-4 py-2 bg-red-600 text-white text-xs font-bold uppercase tracking-wider hover:bg-red-700 transition-colors"
                      >
                        <Trash2 size={14} />
                        <span className="hidden sm:inline">Xóa</span> ({selectedRows.length})
                      </button>
                    )}
                    <button onClick={handleExportCustomers} className="flex items-center gap-2 px-4 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"><Download size={14} /><span className="hidden sm:inline">Xuất Excel</span></button>
                    <button 
                      onClick={() => fileInputRef.current?.click()}
                      className="flex items-center gap-2 px-4 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                    >
                      <FileUp size={14} />
                      <span className="hidden sm:inline">Tải lên file</span>
                    </button>
                    <button 
                      onClick={() => {
                        setEditingId(null);
                        setNewCustomer({ code: '', name: '' });
                        setIsCustomerModalOpen(true);
                      }}
                      className="flex items-center gap-2 px-4 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <Plus size={14} />
                      <span className="hidden sm:inline">Thêm khách hàng</span>
                    </button>
                  </div>
                </div>

                <div className="border border-[#141414] overflow-x-auto max-h-[70vh]">
                  <table className="w-full border-collapse min-w-[600px]">
                    <thead className="sticky top-0 z-10 shadow-sm">
                      <tr className="bg-[#001F3F] text-white text-[11px] uppercase tracking-wider">
                        <th className="border border-[#141414] p-3 w-10">
                          <button onClick={() => {
                            const currentIds = filteredCustomers.map(c => c.id);
                            if (selectedRows.length === currentIds.length) setSelectedRows([]);
                            else setSelectedRows(currentIds);
                          }}>
                            {selectedRows.length > 0 ? <CheckSquare size={14} /> : <Square size={14} />}
                          </button>
                        </th>
                        <th className="border border-[#141414] p-3 text-left">Code</th>
                        <th className="border border-[#141414] p-3 text-left">Name</th>
                        <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                          Thao tác
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredCustomers.map(customer => {
                        const isSelected = selectedRows.includes(customer.id);
                        return (
                          <tr key={customer.id} className={cn(
                            "bg-white text-xs hover:bg-gray-50 transition-colors",
                            isSelected && "bg-blue-50"
                          )}>
                            <td className="border border-[#141414] p-3 text-center">
                              <button onClick={() => toggleRowSelection(customer.id)}>
                                {isSelected ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} />}
                              </button>
                            </td>
                            <td className="border border-[#141414] p-3 font-mono">{customer.code}</td>
                            <td className="border border-[#141414] p-3 font-bold">{customer.name}</td>
                            <td className="border border-[#141414] p-3 text-center">
                              <div className="flex items-center justify-center gap-2">
                                <button 
                                  onClick={() => {
                                    setEditingId(customer.id);
                                    setNewCustomer(customer);
                                    setIsCustomerModalOpen(true);
                                  }}
                                  className="p-1 hover:bg-gray-200 rounded transition-colors"
                                >
                                  <Edit2 size={14} className="text-blue-600" />
                                </button>
                                <button 
                                  onClick={() => {
                                    setDeleteTarget({ id: customer.id, type: 'customer' });
                                    setIsDeleteConfirmOpen(true);
                                  }}
                                  className="p-1 hover:bg-gray-200 rounded transition-colors"
                                >
                                  <Trash2 size={14} className="text-red-600" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </motion.div>
            )}

            {activeTab === 'deliveryNote' && (
              <motion.div 
                key="deliveryNote"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-6"
              >
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 no-print">
                  <div className="flex flex-col sm:flex-row items-start sm:items-center gap-4 sm:gap-6 w-full">
                    <h2 className="font-serif italic text-2xl">Lệnh xuất kho</h2>
                    <div className="flex border-b border-gray-200 w-full sm:w-auto overflow-x-auto">
                      <button
                        onClick={() => setDeliveryNoteSubTab('preview')}
                        className={cn(
                          "px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors border-b-2 whitespace-nowrap",
                          deliveryNoteSubTab === 'preview' ? "border-[#141414] text-[#141414]" : "border-transparent text-gray-400 hover:text-gray-600"
                        )}
                      >
                        Phiếu giao nhận
                      </button>
                      <button
                        onClick={() => setDeliveryNoteSubTab('history')}
                        className={cn(
                          "px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors border-b-2 whitespace-nowrap",
                          deliveryNoteSubTab === 'history' ? "border-[#141414] text-[#141414]" : "border-transparent text-gray-400 hover:text-gray-600"
                        )}
                      >
                        Lịch sử đã Post
                      </button>
                    </div>
                  </div>
                  <div className="flex flex-wrap gap-2 sm:gap-3 w-full sm:w-auto">
                    <div className="relative flex-1 sm:flex-none">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                      <input
                        type="text"
                        placeholder="Tìm kiếm..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="pl-10 pr-4 py-2 border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-full sm:w-64"
                      />
                    </div>
                    {deliveryNoteSubTab === 'preview' && (
                      <div className="flex flex-wrap gap-2 w-full sm:w-auto">
                        <button 
                          onClick={autoAssignLotsFromUI}
                          className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-blue-600 text-white text-[10px] font-bold uppercase tracking-wider hover:bg-blue-700 transition-colors"
                          title="Tự động tìm và gán Lot No cho toàn bộ danh sách hiện tại"
                        >
                          <Package size={14} />
                          <span className="hidden md:inline">Gắn Lot tự động</span>
                        </button>
                        <button 
                          onClick={handlePostDeliveryNote}
                          className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-green-600 text-white text-[10px] font-bold uppercase tracking-wider hover:bg-green-700 transition-colors"
                        >
                          <CheckSquare size={14} />
                          <span className="hidden md:inline">Post xuất kho</span>
                        </button>
                        <button 
                          onClick={() => {
                            window.print();
                          }}
                          className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-3 sm:px-4 py-2 border border-[#141414] text-[10px] font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                        >
                          <Printer size={14} />
                          <span className="hidden md:inline">In phiếu</span>
                        </button>
                        {selectedRows.length > 0 && (
                          <button
                            onClick={() => {
                              setDeleteTarget({ id: 'bulk', type: 'savedDeliveryNote' });
                              setIsDeleteConfirmOpen(true);
                            }}
                            className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-red-600 text-white text-[10px] font-bold uppercase tracking-wider hover:bg-red-700 transition-colors"
                          >
                            <Trash2 size={14} />
                            <span className="hidden md:inline">Xóa</span> ({selectedRows.length})
                          </button>
                        )}
                      </div>
                    )}
                    {deliveryNoteSubTab === 'preview' && (
                      <button 
                        onClick={async () => {
                        const workbook = new ExcelJS.Workbook();
                        const worksheet = workbook.addWorksheet('Delivery Note');

                        worksheet.columns = [
                          { width: 5 },
                          { width: 20 },
                          { width: 20 },
                          { width: 15 },
                          { width: 35 },
                          { width: 8 },
                          { width: 12 },
                          { width: 12 },
                          { width: 15 },
                          { width: 15 },
                          { width: 15 },
                          { width: 15 },
                          { width: 20 },
                          { width: 20 },
                          { width: 10 },
                          { width: 10 },
                          { width: 10 }
                        ];

                        worksheet.columns = [
                          { width: 5 },  // No
                          { width: 15 }, // OVN Sale Order
                          { width: 20 }, // OVN Production Order
                          { width: 12 }, // item
                          { width: 40 }, // Material Name
                          { width: 8 },  // Unit
                          { width: 10 }, // Qty ERP
                          { width: 10 }, // Thực tế
                          { width: 30 }, // Lot No
                          { width: 15 }, // Số lượng thực phát
                          { width: 15 }, // remark
                          { width: 12 }, // Loại chỉ định
                          { width: 10 }, // Brand
                          { width: 25 }, // Customer code
                          { width: 20 }, // Final Destination
                          { width: 10 }, // No.
                          { width: 15 }, // Vị trí
                          { width: 15 }  // STOCK
                        ];

                        const titleRow = worksheet.addRow(['', 'PHIẾU GIAO NHẬN FABRIC']);
                        titleRow.getCell(2).font = { name: 'Times New Roman', size: 18, bold: true };
                        titleRow.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
                        worksheet.mergeCells(1, 2, 1, 15);

                        const subtitleRow = worksheet.addRow(['', 'Delivery Note']);
                        subtitleRow.getCell(2).font = { name: 'Times New Roman', size: 12, italic: true };
                        subtitleRow.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
                        worksheet.mergeCells(2, 2, 2, 15);

                        worksheet.addRow([]);

                        const metaRow1 = worksheet.addRow(['Mã Tài Liệu:', deliveryNoteHeader.documentCode, '', '', '', '', '', '', '', 'Dept:', deliveryNoteHeader.dept]);
                        metaRow1.getCell(1).font = { bold: true };
                        metaRow1.getCell(10).font = { bold: true };
                        worksheet.mergeCells(metaRow1.number, 2, metaRow1.number, 8);
                        worksheet.mergeCells(metaRow1.number, 11, metaRow1.number, 15);

                        const metaRow2 = worksheet.addRow(['TO:', deliveryNoteHeader.to, '', '', '', '', '', '', '', 'Date:', deliveryNoteHeader.date]);
                        metaRow2.getCell(1).font = { bold: true };
                        metaRow2.getCell(10).font = { bold: true };
                        worksheet.mergeCells(metaRow2.number, 2, metaRow2.number, 8);
                        worksheet.mergeCells(metaRow2.number, 11, metaRow2.number, 15);

                        worksheet.addRow([]);

                        const headerRow = worksheet.addRow([
                          'NO', 'OVN SALE ORDER', 'OVN PRODUCTION ORDER', 'ITEM', 'MATERIAL NAME', 
                          'UNIT', 'QTY ERP', 'THỰC TẾ', 'LOT NO', 'SỐ LƯỢNG THỰC PHÁT', 
                          'REMARK', 'LOẠI CHỈ ĐỊNH', 'BRAND', 'CUSTOMER CODE', 'FINAL DESTINATION', 'No.', 'Vị trí', 'STOCK'
                        ]);

                        headerRow.eachCell((cell) => {
                          cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FF001F3F' }
                          };
                          cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 9 };
                          cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                          cell.border = {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                          };
                        });

                        let totalQtyErp = 0;
                        let totalActualQty = 0;
                        let totalActualIssuedQty = 0;
                        const processedGroups = new Set();

                        deliveryNotes.forEach((item) => {
                          totalQtyErp += item.qtyErp;
                          totalActualQty += (item.actualQty || 0);
                          
                          const groupKey = `${item.item}|${item.loaiChiDinh || 'NORMAL'}`;
                          if (!processedGroups.has(groupKey)) {
                            // Sum actualIssuedQty only once per group for the total
                            totalActualIssuedQty += (item.actualIssuedQty || 0);
                            processedGroups.add(groupKey);
                          }

                          const row = worksheet.addRow([
                            item.no,
                            item.ovnSaleOrder,
                            item.ovnProductionOrder,
                            item.item,
                            item.materialName,
                            item.unit,
                            item.qtyErp,
                            item.actualQty,
                            item.lotNo,
                            item.actualIssuedQty,
                            item.remark,
                            item.loaiChiDinh || '',
                            item.brand,
                            item.customerCode,
                            item.finalDestination,
                            item.noCode,
                            item.location,
                            item.stock
                          ]);

                          row.eachCell((cell, colNumber) => {
                            cell.border = {
                              top: { style: 'thin' },
                              left: { style: 'thin' },
                              bottom: { style: 'thin' },
                              right: { style: 'thin' }
                            };
                            cell.font = { size: 9 };
                            
                            // Alignment based on column type
                            if ([1, 4, 6, 9, 12, 13, 16].includes(colNumber)) {
                              cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                            } else if ([7, 8, 10].includes(colNumber)) {
                              cell.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
                              cell.numFmt = '#,##0.000';
                            } else {
                              cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                            }
                          });
                        });

                        let currentGroupKey = '';
                        let startMergeRow = 0;
                        const dataStartRow = 8;

                        deliveryNotes.forEach((item, index) => {
                          const rowIdx = dataStartRow + index;
                          const groupKey = `${item.item}|${item.loaiChiDinh || 'NORMAL'}`;
                          
                          if (groupKey !== currentGroupKey) {
                            if (startMergeRow !== 0 && (rowIdx - 1) > startMergeRow) {
                              // Merge columns that are grouped by Item + loaiChiDinh
                              [4, 5, 6, 9, 10, 11, 12, 13, 14, 15].forEach(col => {
                                worksheet.mergeCells(startMergeRow, col, rowIdx - 1, col);
                              });
                            }
                            currentGroupKey = groupKey;
                            startMergeRow = rowIdx;
                          }
                          
                          if (index === deliveryNotes.length - 1) {
                            if (rowIdx > startMergeRow) {
                              [4, 5, 6, 9, 10, 11, 12, 13, 14, 15].forEach(col => {
                                worksheet.mergeCells(startMergeRow, col, rowIdx, col);
                              });
                            }
                          }
                        });

                        const totalRow = worksheet.addRow([
                          'TỔNG CỘNG', '', '', '', '', '', 
                          totalQtyErp, 
                          totalActualQty, 
                          '', 
                          totalActualIssuedQty,
                          '', '', '', '', '', '', ''
                        ]);
                        worksheet.mergeCells(totalRow.number, 1, totalRow.number, 6);
                        totalRow.eachCell((cell, colNumber) => {
                          cell.font = { bold: true, size: 10 };
                          if ([7, 8, 10].includes(colNumber)) {
                            cell.numFmt = '#,##0.000';
                            cell.alignment = { horizontal: 'right' };
                          }
                          cell.border = {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                          };
                          cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFF5F5F5' }
                          };
                        });
                        totalRow.getCell(1).alignment = { horizontal: 'center' };

                        worksheet.addRow([]);
                        worksheet.addRow([]);
                        const signHeader = worksheet.addRow([
                          'Người lập phiếu (Prepared by)', '', '', '', 
                          'Người nhận hàng (Receiver)', '', '', '', '', '', '', '', 
                          'Thủ kho (Stock keeper)'
                        ]);
                        signHeader.eachCell((cell) => { cell.font = { bold: true, italic: true }; });
                        worksheet.mergeCells(signHeader.number, 1, signHeader.number, 4);
                        worksheet.mergeCells(signHeader.number, 5, signHeader.number, 12);
                        worksheet.mergeCells(signHeader.number, 13, signHeader.number, 17);

                        const signSub = worksheet.addRow([
                          '(Ký, họ tên) (Sign, name)', '', '', '', 
                          '(Ký, họ tên) (Sign, name)', '', '', '', '', '', '', '', 
                          '(Ký, họ tên) (Sign, name)'
                        ]);
                        signSub.eachCell((cell) => { cell.font = { italic: true, size: 9 }; });
                        worksheet.mergeCells(signSub.number, 1, signSub.number, 4);
                        worksheet.mergeCells(signSub.number, 5, signSub.number, 12);
                        worksheet.mergeCells(signSub.number, 13, signSub.number, 17);

                        [signHeader, signSub].forEach(row => {
                          row.eachCell(cell => { cell.alignment = { horizontal: 'center' }; });
                        });

                        const buffer = await workbook.xlsx.writeBuffer();
                        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                        saveAs(blob, "PhieuGiaoNhanFabric.xlsx");
                      }}
                      className="flex items-center gap-2 px-4 py-2 bg-[#141414] text-[#E4E3E0] text-[10px] font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <Download size={14} />
                      Xuất Excel
                    </button>
                    )}
                  </div>
                </div>

                {deliveryNoteSubTab === 'preview' && (
                  <div className="bg-white border border-[#141414] p-4 sm:p-8 shadow-sm print-content">
                  <div className="flex justify-center items-start mb-8">
                    <div className="text-center flex-1">
                      <h1 className="text-lg sm:text-xl font-bold uppercase tracking-widest">Phiếu giao nhận Fabric</h1>
                      <p className="text-xs sm:text-sm italic">Delivery Note</p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-12 gap-y-2 mb-8 text-[10px] sm:text-xs">
                    <div className="grid grid-cols-3 border-b border-gray-200 py-1">
                      <span className="font-bold">Mã Tài Liệu:</span>
                      <input 
                        type="text" 
                        value={deliveryNoteHeader.documentCode} 
                        onChange={(e) => setDeliveryNoteHeader({...deliveryNoteHeader, documentCode: e.target.value})}
                        className="col-span-2 border-none focus:ring-0 p-0 bg-transparent"
                      />
                    </div>
                    <div className="grid grid-cols-3 border-b border-gray-200 py-1">
                      <span className="font-bold">Dept:</span>
                      <input 
                        type="text" 
                        value={deliveryNoteHeader.dept} 
                        onChange={(e) => setDeliveryNoteHeader({...deliveryNoteHeader, dept: e.target.value})}
                        className="col-span-2 border-none focus:ring-0 p-0 bg-transparent"
                      />
                    </div>
                    <div className="grid grid-cols-3 border-b border-gray-200 py-1">
                      <span className="font-bold">TO:</span>
                      <input 
                        type="text" 
                        value={deliveryNoteHeader.to} 
                        onChange={(e) => setDeliveryNoteHeader({...deliveryNoteHeader, to: e.target.value})}
                        className="col-span-2 border-none focus:ring-0 p-0 bg-transparent"
                      />
                    </div>
                    <div className="grid grid-cols-3 border-b border-gray-200 py-1">
                      <span className="font-bold">Date:</span>
                      <input 
                        type="text" 
                        value={deliveryNoteHeader.date} 
                        onChange={(e) => setDeliveryNoteHeader({...deliveryNoteHeader, date: e.target.value})}
                        className="col-span-2 border-none focus:ring-0 p-0 bg-transparent"
                      />
                    </div>
                  </div>

                  <div className="overflow-auto max-h-[65vh] print-table-container">
                    <table className="w-full border-collapse min-w-[1800px]">
                      <thead className="sticky top-0 z-10 shadow-sm">
                        <tr className="bg-[#001F3F] text-white text-[10px] uppercase tracking-wider">
                          <th className="border border-[#141414] p-2 text-center w-10 no-print">
                            <button 
                              onClick={() => {
                                if (selectedRows.length === filteredDeliveryNotes.length) {
                                  setSelectedRows([]);
                                } else {
                                  setSelectedRows(filteredDeliveryNotes.map(i => i.id));
                                }
                              }}
                              className="p-1 hover:bg-white/10 rounded transition-colors"
                            >
                              {selectedRows.length === filteredDeliveryNotes.length && filteredDeliveryNotes.length > 0 ? (
                                <CheckSquare size={14} />
                              ) : (
                                <Square size={14} />
                              )}
                            </button>
                          </th>
                          <th className="border border-[#141414] p-2 text-center w-12">No</th>
                          <th className="border border-[#141414] p-2 text-left">OVN Sale Order</th>
                          <th className="border border-[#141414] p-2 text-left">OVN Production Order</th>
                          <th className="border border-[#141414] p-2 text-left">item</th>
                          <th className="border border-[#141414] p-2 text-left">Material Name</th>
                          <th className="border border-[#141414] p-2 text-center">Unit</th>
                          <th className="border border-[#141414] p-2 text-right">Qty ERP</th>
                          <th className="border border-[#141414] p-2 text-right">Thực tế</th>
                          <th className="border border-[#141414] p-2 text-left">Lot No</th>
                          <th className="border border-[#141414] p-2 text-right">Số lượng thực phát</th>
                          <th className="border border-[#141414] p-2 text-left">remark</th>
                          <th className="border border-[#141414] p-2 text-left">Loại chỉ định</th>
                          <th className="border border-[#141414] p-2 text-left">Brand</th>
                          <th className="border border-[#141414] p-2 text-left">Customer code</th>
                          <th className="border border-[#141414] p-2 text-left">Final Destination</th>
                          <th className="border border-[#141414] p-2 text-left">No.</th>
                          <th className="border border-[#141414] p-2 text-left">Vị trí</th>
                          <th className="border border-[#141414] p-2 text-left">STOCK</th>
                          <th className="border border-[#141414] p-2 no-print">Thao tác</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredDeliveryNotes.length === 0 ? (
                          <tr>
                            <td colSpan={19} className="p-12 text-center text-gray-400 italic">
                              Chưa có dữ liệu. Vui lòng tải lên file nguồn để cập nhật.
                            </td>
                          </tr>
                        ) : (
                          filteredDeliveryNotes.map((item, index) => {
                            const isDesignationMatch = (remark: string, prodOrder: string, saleOrder: string, noCode: string) => {
                              if (!remark) return true;
                              // Remove all whitespace and split by /
                              const codes = remark.replace(/\s+/g, '').toLowerCase().split('/');
                              const targetRpro = prodOrder.replace(/\s+/g, '').toLowerCase();
                              const targetSo = saleOrder.replace(/\s+/g, '').toLowerCase();
                              const targetNo = noCode.replace(/\s+/g, '').toLowerCase();
                              
                              return (targetRpro && codes.includes(targetRpro)) || 
                                     (targetSo && codes.includes(targetSo)) || 
                                     (targetNo && codes.includes(targetNo));
                            };

                            const hasMismatch = !isDesignationMatch(item.remark, item.ovnProductionOrder, item.ovnSaleOrder, item.noCode);

                            // RowSpan logic: Group by ITEM and loaiChiDinh
                            const isFirstInItemGroup = index === 0 || 
                              filteredDeliveryNotes[index - 1].item !== item.item || 
                              filteredDeliveryNotes[index - 1].loaiChiDinh !== item.loaiChiDinh;
                            let itemGroupSize = 0;
                            if (isFirstInItemGroup) {
                              for (let i = index; i < filteredDeliveryNotes.length; i++) {
                                if (filteredDeliveryNotes[i].item === item.item && filteredDeliveryNotes[i].loaiChiDinh === item.loaiChiDinh) {
                                  itemGroupSize++;
                                } else {
                                  break;
                                }
                              }
                            }

                            // Aggregate lots for the group
                            const getGroupedLots = (itemCode: string, startIndex: number, size: number) => {
                              const groupItems = filteredDeliveryNotes.slice(startIndex, startIndex + size);
                              const allLots: {lotNo: string, qty: number, stock: string, location: string, loaiChiDinh: string}[] = [];
                              
                              groupItems.forEach(gi => {
                                if (gi.assignedLots) {
                                  gi.assignedLots.forEach(lot => {
                                    const existing = allLots.find(l => l.lotNo === lot.lotNo);
                                    if (existing) {
                                      existing.qty += lot.qty;
                                    } else {
                                      allLots.push({ ...lot });
                                    }
                                  });
                                }
                              });
                              return allLots;
                            };

                            const groupLots = isFirstInItemGroup ? getGroupedLots(item.item, index, itemGroupSize) : [];

                            return (
                              <tr 
                                key={item.id} 
                                className={cn(
                                  "text-[11px] transition-colors",
                                  !item.noCode ? "bg-yellow-200 text-red-600 font-bold" : "bg-white hover:bg-gray-50",
                                  hasMismatch ? "bg-red-100 text-red-700" : "",
                                  selectedRows.includes(item.id) && "bg-blue-50"
                                )}
                              >
                                <td className="border border-[#141414] p-2 text-center no-print">
                                  <button 
                                    onClick={() => {
                                      if (selectedRows.includes(item.id)) {
                                        setSelectedRows(selectedRows.filter(id => id !== item.id));
                                      } else {
                                        setSelectedRows([...selectedRows, item.id]);
                                      }
                                    }}
                                    className="p-1 hover:bg-gray-100 rounded transition-colors"
                                  >
                                    {selectedRows.includes(item.id) ? (
                                      <CheckSquare size={14} className="text-[#141414]" />
                                    ) : (
                                      <Square size={14} className="text-gray-400" />
                                    )}
                                  </button>
                                </td>
                                <td className="border border-[#141414] p-2 text-center font-bold">{item.no}</td>
                                <td className="border border-[#141414] p-2">{item.ovnSaleOrder}</td>
                                <td className={cn(
                                  "border border-[#141414] p-2",
                                  hasMismatch && "font-bold underline decoration-double"
                                )}>
                                  {item.ovnProductionOrder}
                                </td>
                                <td className="border border-[#141414] p-2 font-mono">{item.item}</td>
                                <td className="border border-[#141414] p-2">{item.materialName}</td>
                                <td className="border border-[#141414] p-2 text-center">{item.unit}</td>
                                <td className="border border-[#141414] p-2 text-right font-mono">{item.qtyErp.toLocaleString()}</td>
                                
                                <td className="border border-[#141414] p-2 text-right font-mono bg-blue-50/30">
                                  <input 
                                    type="number"
                                    value={item.actualQty || 0}
                                    onChange={(e) => handleEditDeliveryNoteItem(index, 'actualQty', Number(e.target.value))}
                                    className="w-full bg-transparent text-right focus:outline-none focus:ring-1 focus:ring-[#141414]"
                                  />
                                </td>
                                
                                {isFirstInItemGroup ? (
                                  <>
                                    <td rowSpan={itemGroupSize} className="border border-[#141414] p-2 bg-blue-50/30 align-middle">
                                      {groupLots.length > 0 ? (
                                        <div className="flex flex-col gap-1">
                                          {groupLots.map((lot, idx) => (
                                            <div key={idx} className={cn(idx > 0 && "border-t border-gray-200 pt-1")}>
                                              {lot.lotNo} ({lot.qty.toLocaleString()})
                                            </div>
                                          ))}
                                        </div>
                                      ) : (
                                        item.lotNo
                                      )}
                                    </td>
                                    <td rowSpan={itemGroupSize} className="border border-[#141414] p-2 text-right font-mono align-middle bg-gray-50/50 text-blue-600 font-bold">
                                      {groupLots.length > 0 ? (
                                        groupLots.reduce((sum, lot) => sum + lot.qty, 0).toLocaleString()
                                      ) : (
                                        filteredDeliveryNotes.slice(index, index + itemGroupSize)
                                          .reduce((sum, gi) => sum + (gi.actualQty || 0), 0).toLocaleString()
                                      )}
                                    </td>
                                  </>
                                ) : null}

                                <td className={cn(
                                  "border border-[#141414] p-2",
                                  hasMismatch && "bg-red-200 font-bold"
                                )}>
                                  <input 
                                    type="text"
                                    value={item.remark}
                                    onChange={(e) => handleEditDeliveryNoteItem(index, 'remark', e.target.value)}
                                    className="w-full bg-transparent focus:outline-none focus:ring-1 focus:ring-[#141414]"
                                  />
                                  {hasMismatch && (
                                    <div className="text-[9px] mt-1 flex items-center gap-1">
                                      <AlertTriangle size={10} />
                                      <span>Sai mã chỉ định</span>
                                    </div>
                                  )}
                                </td>
                                <td className="border border-[#141414] p-2 bg-blue-50/30">
                                  {item.assignedLots && item.assignedLots.length > 0 ? (
                                    <div className="flex flex-col gap-1">
                                      {item.assignedLots.map((lot, idx) => (
                                        <div key={idx} className={cn(idx > 0 && "border-t border-gray-200 pt-1")}>
                                          {lot.loaiChiDinh}
                                        </div>
                                      ))}
                                    </div>
                                  ) : (
                                    item.loaiChiDinh
                                  )}
                                </td>
                                <td className="border border-[#141414] p-2">{item.brand}</td>
                                <td className="border border-[#141414] p-2">{item.customerCode}</td>
                                <td className="border border-[#141414] p-2">{item.finalDestination}</td>
                                <td className="border border-[#141414] p-2">{item.noCode}</td>
                                <td className="border border-[#141414] p-2">
                                  {item.assignedLots && item.assignedLots.length > 0 ? (
                                    <div className="flex flex-col gap-1">
                                      {item.assignedLots.map((lot, idx) => (
                                        <div key={idx} className={cn(idx > 0 && "border-t border-gray-200 pt-1")}>
                                          {getLocationByItemAndLot(item.item, lot.lotNo)}
                                        </div>
                                      ))}
                                    </div>
                                  ) : (
                                    getLocationByItemAndLot(item.item, item.lotNo)
                                  )}
                                </td>
                                {isFirstInItemGroup ? (() => {
                                  const groupItems = filteredDeliveryNotes.slice(index, index + itemGroupSize);
                                  const analysis = analyzeGroupStock(item.item, groupItems);
                                  
                                  return (
                                    <td rowSpan={itemGroupSize} className={cn(
                                      "border border-[#141414] p-2 bg-blue-50/30 align-middle",
                                      analysis.detail === 'Không có tồn' ? "bg-red-500 text-white font-bold" : ""
                                    )}>
                                      <div className="flex flex-col gap-1">
                                        <div className="font-medium">{analysis.detail}</div>
                                        {analysis.shortage > 0 && (
                                          <div className="text-[10px] text-red-600 font-bold bg-red-50 px-1 rounded border border-red-200 w-fit">
                                            THIẾU: {analysis.shortage.toLocaleString()}
                                          </div>
                                        )}
                                        {analysis.shortage === 0 && groupLots.length > 0 && (
                                          <div className="text-[10px] text-green-600 font-bold bg-green-50 px-1 rounded border border-green-200 w-fit">
                                            ĐỦ TỒN
                                          </div>
                                        )}
                                      </div>
                                    </td>
                                  );
                                })() : null}
                                <td className="border border-[#141414] p-2 no-print text-center">
                                  <div className="flex items-center justify-center gap-2">
                                    <button 
                                      onClick={() => handleEditDeliveryNoteItemClick(item.id)}
                                      className="text-blue-600 hover:text-blue-800 transition-colors"
                                      title="Sửa"
                                    >
                                      <Edit2 size={14} />
                                    </button>
                                    <button 
                                      onClick={() => handleDeleteDeliveryNoteItem(item.id)}
                                      className="text-red-600 hover:text-red-800 transition-colors"
                                      title="Xóa"
                                    >
                                      <Trash2 size={14} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                          })
                        )}
                        {filteredDeliveryNotes.length > 0 && (
                          <tr className="bg-gray-50 font-bold text-[11px]">
                            <td colSpan={6} className="border border-[#141414] p-2 text-center uppercase">Tổng cộng</td>
                            <td className="border border-[#141414] p-2 text-right font-mono">
                              {filteredDeliveryNotes.reduce((sum, item) => sum + item.qtyErp, 0).toLocaleString()}
                            </td>
                            <td className="border border-[#141414] p-2 text-right font-mono">
                              {filteredDeliveryNotes.reduce((sum, item) => sum + (item.actualQty || 0), 0).toLocaleString()}
                            </td>
                            <td className="border border-[#141414] p-2"></td>
                            <td className="border border-[#141414] p-2 text-right font-mono">
                              {Array.from(new Set(filteredDeliveryNotes.map(item => item.item)))
                                .reduce((sum, itemCode) => {
                                  const item = filteredDeliveryNotes.find(dn => dn.item === itemCode);
                                  return sum + (item?.actualIssuedQty || 0);
                                }, 0).toLocaleString()}
                            </td>
                            <td colSpan={7} className="border border-[#141414] p-2"></td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>

                  <div className="grid grid-cols-3 gap-8 mt-12 text-center text-xs italic">
                    <div>
                      <p className="font-bold mb-12">Người lập phiếu (Prepared by)</p>
                      <p>(Ký, họ tên) (Sign, name)</p>
                    </div>
                    <div>
                      <p className="font-bold mb-12">Người nhận hàng (Receiver)</p>
                      <p>(Ký, họ tên) (Sign, name)</p>
                    </div>
                    <div>
                      <p className="font-bold mb-12">Thủ kho (Stock keeper)</p>
                      <p>(Ký, họ tên) (Sign, name)</p>
                    </div>
                  </div>
                </div>

                )}
                
                {deliveryNoteSubTab === 'history' && (
                  <div className="space-y-4">
                    {savedDeliveryNotes.length === 0 ? (
                      <div className="text-center py-12 bg-white border border-dashed border-gray-300">
                        <p className="text-gray-400 italic">Chưa có phiếu nào được Post</p>
                      </div>
                    ) : (
                      savedDeliveryNotes.map((note) => (
                        <div key={note.id} className="border border-[#141414] p-6 bg-white flex justify-between items-center hover:shadow-md transition-shadow">
                          <div className="flex items-center gap-6">
                            <div className="w-12 h-12 bg-[#141414] text-[#E4E3E0] flex items-center justify-center">
                              <FileText size={20} />
                            </div>
                            <div>
                              <p className="font-bold text-lg">Phiếu ngày: {note.date}</p>
                              <div className="flex gap-4 text-xs opacity-60 mt-1">
                                <span>Số lượng dòng: {note.items.length}</span>
                                <span>ID: {note.id.slice(0, 8)}...</span>
                              </div>
                            </div>
                          </div>
                          <div className="flex gap-4">
                            <button 
                              onClick={() => {
                                setDeliveryNotes(note.items);
                                setDeliveryNoteSubTab('preview');
                              }}
                              className="px-4 py-2 border border-[#141414] text-[10px] font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                            >
                              Xem lại
                            </button>
                            <button 
                              onClick={() => {
                                setDeleteTarget({ id: note.id, type: 'savedDeliveryNote' });
                                setIsDeleteConfirmOpen(true);
                              }}
                              className="px-4 py-2 border border-red-600 text-red-600 text-[10px] font-bold uppercase tracking-wider hover:bg-red-600 hover:text-white transition-colors"
                            >
                              Xóa
                            </button>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                )}
              </motion.div>
            )}
            {activeTab === 'location' && (
              <motion.div 
                key="location"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-6"
              >
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 no-print">
                  <div className="flex flex-col sm:flex-row items-start sm:items-center gap-4 sm:gap-6 w-full">
                    <h2 className="font-serif italic text-2xl">Quản lý Vị Trí</h2>
                    <div className="flex border-b border-gray-200 w-full sm:w-auto overflow-x-auto">
                      <button
                        onClick={() => {
                          setLocationSubTab('input');
                          setScanMode('INPUT');
                        }}
                        className={cn(
                          "px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors border-b-2 whitespace-nowrap",
                          locationSubTab === 'input' ? "border-[#141414] text-[#141414]" : "border-transparent text-gray-400 hover:text-gray-600"
                        )}
                      >
                        NHẬP VỊ TRÍ
                      </button>
                      <button
                        onClick={() => {
                          setLocationSubTab('output');
                          setScanMode('OUTPUT');
                        }}
                        className={cn(
                          "px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors border-b-2 whitespace-nowrap",
                          locationSubTab === 'output' ? "border-[#141414] text-[#141414]" : "border-transparent text-gray-400 hover:text-gray-600"
                        )}
                      >
                        XUẤT VỊ TRÍ
                      </button>
                      <button
                        onClick={() => setLocationSubTab('inventory')}
                        className={cn(
                          "px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors border-b-2 whitespace-nowrap",
                          locationSubTab === 'inventory' ? "border-[#141414] text-[#141414]" : "border-transparent text-gray-400 hover:text-gray-600"
                        )}
                      >
                        TỒN VỊ TRÍ
                      </button>
                    </div>
                    <div className="relative w-full sm:w-auto">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                      <input 
                        type="text" 
                        placeholder="Tìm kiếm..."
                        value={locationSearch}
                        onChange={(e) => setLocationSearch(e.target.value)}
                        className="pl-9 pr-4 py-2 bg-white border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-full sm:w-64"
                      />
                    </div>
                  </div>
                  <div className="flex flex-wrap gap-2 sm:gap-3 w-full sm:w-auto">
                    <button 
                      onClick={() => {
                        setEditingId(null);
                        setNewLocationEntry({ qrcode: '', sku: '', partner: '', date: '', location: '', note: '', quantity: 1 });
                        if (locationSubTab === 'output') {
                          setScanMode('OUTPUT');
                        } else {
                          setScanMode('INPUT');
                        }
                        setIsLocationModalOpen(true);
                      }}
                      className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <Plus size={14} />
                      <span className="hidden sm:inline">Thêm vị trí</span>
                    </button>
                    <button onClick={handleExportLocations} className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"><Download size={14} /><span className="hidden sm:inline">Xuất Excel</span></button>
                    <button 
                      onClick={() => fileInputRef.current?.click()}
                      className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <FileUp size={14} />
                      <span className="hidden sm:inline">Nhập Excel</span>
                    </button>
                    {selectedRows.length > 0 && (
                      <button 
                        onClick={handleBulkDelete}
                        className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2 bg-red-600 text-white text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                      >
                        <Trash2 size={14} />
                        Xóa ({selectedRows.length})
                      </button>
                    )}
                  </div>
                </div>

                {(locationSubTab === 'input' || locationSubTab === 'output') && (
                  <div className="bg-white border border-[#141414] p-4 space-y-4 no-print">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                      <h3 className="text-xs font-bold uppercase tracking-widest flex items-center gap-2">
                        <MapPin size={14} />
                        Scan / Dán dữ liệu nhanh
                      </h3>
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => setScanMode('INPUT')}
                          className={`flex-1 sm:flex-none px-6 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border border-[#141414] ${
                            scanMode === 'INPUT' 
                              ? 'bg-blue-600 text-white opacity-100' 
                              : 'bg-white text-blue-600 opacity-30 hover:opacity-50'
                          }`}
                        >
                          NHẬP
                        </button>
                        <button
                          onClick={() => setScanMode('OUTPUT')}
                          className={`flex-1 sm:flex-none px-6 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border border-[#141414] ${
                            scanMode === 'OUTPUT' 
                              ? 'bg-red-600 text-white opacity-100' 
                              : 'bg-white text-red-600 opacity-30 hover:opacity-50'
                          }`}
                        >
                          XUẤT
                        </button>
                      </div>
                    </div>
                    <div className="flex flex-col sm:flex-row gap-4">
                      <textarea 
                        value={scanInput}
                        onChange={(e) => setScanInput(e.target.value)}
                        placeholder="Dán mã tại đây. Vị trí bắt đầu bằng FB, QRCODE bắt đầu bằng AWB..."
                        className="flex-1 h-24 p-3 border border-[#141414] text-xs font-mono focus:outline-none focus:ring-1 focus:ring-[#141414] resize-none"
                      />
                      <button 
                        onClick={handleProcessScanInput}
                        disabled={!scanInput.trim()}
                        className="py-4 sm:py-0 sm:px-8 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity disabled:opacity-50 flex flex-row sm:flex-col items-center justify-center gap-2"
                      >
                        <Save size={20} />
                        Cập nhật
                      </button>
                    </div>
                  </div>
                )}

                {(locationSubTab === 'input' || locationSubTab === 'output') && (
                  <div className="space-y-4">
                    <div className="border border-[#141414] overflow-auto max-h-[70vh]">
                    <table className="w-full border-collapse">
                      <thead className="sticky top-0 z-10 shadow-sm">
                        <tr className="bg-[#001F3F] text-white text-[11px] uppercase tracking-wider">
                          <th className="border border-[#141414] p-3 w-10">
                            <button onClick={() => {
                              const currentIds = filteredLocationEntries.map(i => i.id);
                              if (selectedRows.length === currentIds.length) setSelectedRows([]);
                              else setSelectedRows(currentIds);
                            }}>
                              {selectedRows.length > 0 ? <CheckSquare size={14} /> : <Square size={14} />}
                            </button>
                          </th>
                          <th className="border border-[#141414] p-3 text-left">QRCODE</th>
                          <th className="border border-[#141414] p-3 text-left">Mã</th>
                          <th className="border border-[#141414] p-3 text-left">NCC</th>
                          <th className="border border-[#141414] p-3 text-left">NGÀY</th>
                          <th className="border border-[#141414] p-3 text-left">Vị trí</th>
                          <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                            Thao tác
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredLocationEntries.length === 0 ? (
                          <tr>
                            <td colSpan={6} className="border border-[#141414] p-8 text-center italic text-gray-400">
                              Chưa có dữ liệu vị trí. Hãy nhập file Excel.
                            </td>
                          </tr>
                        ) : (
                          (() => {
                            const rows: React.ReactNode[] = [];
                            let lastScanType: string | undefined = undefined;
                            
                            const sortedEntries = [...filteredLocationEntries].reverse();

                            sortedEntries.forEach((entry, index) => {
                              if (entry.scanType && entry.scanType !== lastScanType) {
                                rows.push(
                                  <tr key={`header-${entry.scanType}-${index}`} className={cn(
                                    "font-bold text-sm",
                                    entry.scanType === 'INPUT' ? "bg-yellow-400" : "bg-cyan-400"
                                  )}>
                                    <td colSpan={6} className="border border-[#141414] p-2 uppercase">
                                      {entry.scanType}
                                    </td>
                                  </tr>
                                );
                                lastScanType = entry.scanType;
                              }
                              
                              rows.push(
                                <tr key={entry.id} className={cn(
                                  "bg-white text-xs hover:bg-gray-50 transition-colors",
                                  selectedRows.includes(entry.id) && "bg-blue-50"
                                )}>
                                  <td className="border border-[#141414] p-3 text-center">
                                    <button onClick={() => toggleRowSelection(entry.id)}>
                                      {selectedRows.includes(entry.id) ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} />}
                                    </button>
                                  </td>
                                  <td className="border border-[#141414] p-3 font-mono">{entry.qrcode}</td>
                                  <td className="border border-[#141414] p-3 font-bold">{entry.sku}</td>
                                  <td className="border border-[#141414] p-3">{entry.partner}</td>
                                  <td className="border border-[#141414] p-3">{entry.date}</td>
                                  <td className="border border-[#141414] p-3 font-bold text-blue-600">{entry.location}</td>
                                  <td className="border border-[#141414] p-3 text-center">
                                    <div className="flex items-center justify-center gap-2">
                                      <button 
                                        onClick={() => {
                                          setEditingId(entry.id);
                                          setNewLocationEntry({ ...entry });
                                          setIsLocationModalOpen(true);
                                        }}
                                        className="p-1 hover:bg-gray-200 rounded transition-colors text-blue-600"
                                      >
                                        <Edit2 size={14} />
                                      </button>
                                      <button 
                                        onClick={() => {
                                          setDeleteTarget({ id: entry.id, type: 'location' });
                                          setIsDeleteConfirmOpen(true);
                                        }}
                                        className="p-1 hover:bg-gray-200 rounded transition-colors text-red-600"
                                      >
                                        <Trash2 size={14} />
                                      </button>
                                    </div>
                                  </td>
                                </tr>
                              );
                            });
                            
                            return rows.reverse();
                          })()
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
                )}

                {locationSubTab === 'inventory' && (
                  <div className="border border-[#141414] overflow-auto max-h-[70vh]">
                    <table className="w-full border-collapse">
                      <thead className="sticky top-0 z-10 shadow-sm">
                        <tr className="bg-[#001F3F] text-white text-[11px] uppercase tracking-wider">
                          <th className="border border-[#141414] p-3 w-10">
                            <button onClick={() => {
                              const currentIds = filteredLocationInventoryEntries.map(i => i.id);
                              if (selectedRows.length === currentIds.length) setSelectedRows([]);
                              else setSelectedRows(currentIds);
                            }}>
                              {selectedRows.length > 0 ? <CheckSquare size={14} /> : <Square size={14} />}
                            </button>
                          </th>
                          <th className="border border-[#141414] p-3 text-left">QRCODE</th>
                          <th className="border border-[#141414] p-3 text-left">Mã</th>
                          <th className="border border-[#141414] p-3 text-left">NCC</th>
                          <th className="border border-[#141414] p-3 text-left">NGÀY</th>
                          <th className="border border-[#141414] p-3 text-left">Vị trí</th>
                          <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                            Thao tác
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredLocationInventoryEntries.length === 0 ? (
                          <tr>
                            <td colSpan={6} className="border border-[#141414] p-8 text-center italic text-gray-400">
                              Không có dữ liệu tồn theo vị trí. Hãy nhập file Excel (Sheet 3).
                            </td>
                          </tr>
                        ) : (
                          filteredLocationInventoryEntries.map(entry => {
                            const isSelected = selectedRows.includes(entry.id);
                            return (
                              <tr key={entry.id} className={cn(
                                "bg-white text-xs hover:bg-gray-50 transition-colors",
                                isSelected && "bg-blue-50"
                              )}>
                                <td className="border border-[#141414] p-3 text-center">
                                  <button onClick={() => toggleRowSelection(entry.id)}>
                                    {isSelected ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} />}
                                  </button>
                                </td>
                                <td className="border border-[#141414] p-3 font-mono">{entry.qrcode}</td>
                              <td className="border border-[#141414] p-3 font-bold">{entry.sku}</td>
                              <td className="border border-[#141414] p-3">{entry.partner}</td>
                              <td className="border border-[#141414] p-3">{entry.date}</td>
                              <td className="border border-[#141414] p-3 font-bold text-blue-600">{entry.location}</td>
                              <td className="border border-[#141414] p-3 text-center">
                                <div className="flex items-center justify-center gap-2">
                                  <button 
                                    onClick={() => {
                                      setEditingId(entry.id);
                                      setNewLocationEntry({ ...entry });
                                      setIsLocationModalOpen(true);
                                    }}
                                    className="p-1 hover:bg-gray-200 rounded transition-colors text-blue-600"
                                  >
                                    <Edit2 size={14} />
                                  </button>
                                  <button 
                                    onClick={() => {
                                      setDeleteTarget({ id: entry.id, type: 'location', qrcode: entry.qrcode });
                                      setIsDeleteConfirmOpen(true);
                                    }}
                                    className="p-1 hover:bg-gray-200 rounded transition-colors text-red-600"
                                  >
                                    <Trash2 size={14} />
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                    </table>
                  </div>
                )}
              </motion.div>
            )}
          </AnimatePresence>

      <AnimatePresence>
        {isLocationModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-[#E4E3E0] border border-[#141414] p-8 w-full max-w-md space-y-6"
            >
              <h3 className="font-serif italic text-2xl">
                {editingId ? 'Chỉnh sửa vị trí' : 'Thêm vị trí mới'}
              </h3>
              <form onSubmit={handleLocationSubmit} className="space-y-4">
                <div className="flex items-center gap-2 p-1 bg-white/50 border border-[#141414]/10 rounded">
                  <button 
                    type="button"
                    onClick={() => setScanMode('INPUT')}
                    className={cn(
                      "flex-1 py-2 text-[10px] font-bold uppercase tracking-wider transition-all",
                      scanMode === 'INPUT' 
                        ? "bg-green-500 text-white shadow-inner" 
                        : "text-gray-400 hover:bg-gray-100"
                    )}
                  >
                    NHẬP
                  </button>
                  <button 
                    type="button"
                    onClick={() => setScanMode('OUTPUT')}
                    className={cn(
                      "flex-1 py-2 text-[10px] font-bold uppercase tracking-wider transition-all",
                      scanMode === 'OUTPUT' 
                        ? "bg-red-500 text-white shadow-inner" 
                        : "text-gray-400 hover:bg-gray-100"
                    )}
                  >
                    XUẤT
                  </button>
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] uppercase font-bold opacity-50">QR CODE</label>
                  <input 
                    type="text" 
                    required
                    value={newLocationEntry.qrcode} 
                    onChange={(e) => {
                      const qrcode = e.target.value;
                      const parsed = parseQRCode(qrcode);
                      setNewLocationEntry({
                        ...newLocationEntry,
                        qrcode,
                        sku: parsed?.sku || newLocationEntry.sku,
                        partner: parsed?.partner || newLocationEntry.partner,
                        date: parsed?.date || ''
                      });
                    }}
                    className="w-full bg-transparent border-b border-[#141414] py-2 focus:outline-none focus:border-blue-500 transition-colors font-mono"
                    placeholder="Nhập hoặc dán mã QR..."
                  />
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Mã Hàng</label>
                    <input 
                      type="text" 
                      required
                      value={newLocationEntry.sku} 
                      onChange={(e) => setNewLocationEntry({...newLocationEntry, sku: e.target.value})}
                      className="w-full bg-transparent border-b border-[#141414] py-2 focus:outline-none focus:border-blue-500 transition-colors"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">NCC</label>
                    <input 
                      type="text" 
                      required
                      value={newLocationEntry.partner} 
                      onChange={(e) => setNewLocationEntry({...newLocationEntry, partner: e.target.value})}
                      className="w-full bg-transparent border-b border-[#141414] py-2 focus:outline-none focus:border-blue-500 transition-colors"
                    />
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Ngày</label>
                    <input 
                      type="text" 
                      required
                      value={newLocationEntry.date} 
                      onChange={(e) => setNewLocationEntry({...newLocationEntry, date: e.target.value})}
                      className="w-full bg-transparent border-b border-[#141414] py-2 focus:outline-none focus:border-blue-500 transition-colors"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Vị trí</label>
                    <input 
                      type="text" 
                      required
                      value={newLocationEntry.location} 
                      onChange={(e) => setNewLocationEntry({...newLocationEntry, location: e.target.value})}
                      className="w-full bg-transparent border-b border-[#141414] py-2 focus:outline-none focus:border-blue-500 transition-colors font-bold text-blue-600"
                    />
                  </div>
                </div>
                <div className="flex gap-4 pt-4">
                  <button 
                    type="button"
                    onClick={() => setIsLocationModalOpen(false)}
                    className="flex-1 py-3 border border-[#141414] text-[10px] uppercase font-bold tracking-widest hover:bg-white transition-colors"
                  >
                    Hủy
                  </button>
                  <button 
                    type="submit"
                    className="flex-1 py-3 bg-[#141414] text-[#E4E3E0] text-[10px] uppercase font-bold tracking-widest hover:opacity-90 transition-opacity"
                  >
                    {editingId ? 'Cập nhật' : 'Thêm mới'}
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {isProductModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-[#E4E3E0] border border-[#141414] p-8 w-full max-w-md space-y-6"
            >
              <h3 className="font-serif italic text-2xl">
                {editingId ? 'Chỉnh sửa sản phẩm' : 'Thêm sản phẩm mới'}
              </h3>
              <form onSubmit={handleAddProduct} className="space-y-4">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Mã Hàng (SKU)</label>
                    <input 
                      required
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={newProduct.sku}
                      onChange={e => setNewProduct({...newProduct, sku: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Đơn vị</label>
                    <input 
                      required
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={newProduct.unit}
                      onChange={e => setNewProduct({...newProduct, unit: e.target.value})}
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] uppercase font-bold opacity-50">Tên hàng</label>
                  <input 
                    required
                    className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                    value={newProduct.name}
                    onChange={e => setNewProduct({...newProduct, name: e.target.value})}
                  />
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Lot no</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={newProduct.lotNo}
                      onChange={e => setNewProduct({...newProduct, lotNo: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Ghi chú</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={newProduct.ghiChu}
                      onChange={e => setNewProduct({...newProduct, ghiChu: e.target.value})}
                    />
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Mã chỉ định</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={newProduct.designationCode}
                      onChange={e => setNewProduct({...newProduct, designationCode: e.target.value.replace(/\s+/g, '')})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Loại chỉ định</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={newProduct.loaiChiDinh}
                      onChange={e => setNewProduct({...newProduct, loaiChiDinh: e.target.value})}
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] uppercase font-bold opacity-50">Số lượng</label>
                  <input 
                    type="number"
                    step="any"
                    className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none font-bold text-blue-600"
                    value={newProduct.quantity || 0}
                    onChange={e => setNewProduct({...newProduct, quantity: Number(e.target.value)})}
                  />
                </div>
                <div className="flex gap-4 pt-4">
                  <button 
                    type="button"
                    onClick={() => setIsProductModalOpen(false)}
                    className="flex-1 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                  >
                    Hủy
                  </button>
                  <button 
                    type="submit"
                    className="flex-1 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                  >
                    Lưu
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {isTransactionModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-[#E4E3E0] border border-[#141414] p-8 w-full max-w-md space-y-6"
            >
              <h3 className="font-serif italic text-2xl">
                {editingId ? 'Chỉnh sửa giao dịch' : (activeTab === 'inbound' ? 'Cập nhật Nhập kho' : 'Cập nhật Xuất kho')}
              </h3>
              <form onSubmit={handleAddTransaction} className="space-y-4">
                <div className="space-y-1">
                  <label className="text-[10px] uppercase font-bold opacity-50">Sản phẩm</label>
                  <select 
                    required
                    className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                    value={newTransaction.productId}
                    onChange={e => {
                      const p = products.find(prod => prod.id === e.target.value);
                      setNewTransaction({
                        ...newTransaction, 
                        productId: e.target.value,
                        lotNo: p?.lotNo,
                        ghiChu: p?.ghiChu,
                        designationCode: p?.designationCode
                      });
                    }}
                  >
                    <option value="">Chọn sản phẩm</option>
                    {products.map(p => (
                      <option key={p.id} value={p.id}>{p.sku} - {p.name}</option>
                    ))}
                  </select>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Số lượng</label>
                    <input 
                      required
                      type="number"
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none font-mono"
                      value={newTransaction.quantity || ''}
                      onChange={e => setNewTransaction({...newTransaction, quantity: Number(e.target.value)})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Ngày (dd/mm/yyyy)</label>
                    <input 
                      required
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none font-mono"
                      value={newTransaction.date}
                      onChange={e => setNewTransaction({...newTransaction, date: e.target.value})}
                    />
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Lot No</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none font-mono"
                      value={newTransaction.lotNo || ''}
                      onChange={e => setNewTransaction({...newTransaction, lotNo: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Mã chỉ định</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none font-mono"
                      value={newTransaction.designationCode || ''}
                      onChange={e => setNewTransaction({...newTransaction, designationCode: e.target.value})}
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] uppercase font-bold opacity-50">Loại chỉ định</label>
                  <input 
                    className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                    value={newTransaction.loaiChiDinh}
                    onChange={e => setNewTransaction({...newTransaction, loaiChiDinh: e.target.value})}
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] uppercase font-bold opacity-50">Ghi chú</label>
                  <input 
                    className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                    value={newTransaction.ghiChu}
                    onChange={e => setNewTransaction({...newTransaction, ghiChu: e.target.value})}
                  />
                </div>
                <div className="flex gap-4 pt-4">
                  <button 
                    type="button"
                    onClick={() => {
                      setIsTransactionModalOpen(false);
                      setEditingId(null);
                    }}
                    className="flex-1 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                  >
                    Hủy
                  </button>
                  <button 
                    type="submit"
                    className="flex-1 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                  >
                    Cập nhật
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {isCustomerModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-[#E4E3E0] border border-[#141414] p-8 w-full max-w-md space-y-6"
            >
              <h3 className="font-serif italic text-2xl">
                {editingId ? 'Chỉnh sửa khách hàng' : 'Thêm khách hàng mới'}
              </h3>
              <form onSubmit={handleAddCustomer} className="space-y-4">
                <div className="space-y-1">
                  <label className="text-[10px] uppercase font-bold opacity-50">Mã khách hàng</label>
                  <input 
                    required
                    className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                    value={newCustomer.code}
                    onChange={e => setNewCustomer({...newCustomer, code: e.target.value})}
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] uppercase font-bold opacity-50">Tên khách hàng</label>
                  <input 
                    required
                    className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                    value={newCustomer.name}
                    onChange={e => setNewCustomer({...newCustomer, name: e.target.value})}
                  />
                </div>
                <div className="flex gap-4 pt-4">
                  <button 
                    type="button"
                    onClick={() => setIsCustomerModalOpen(false)}
                    className="flex-1 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                  >
                    Hủy
                  </button>
                  <button 
                    type="submit"
                    className="flex-1 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                  >
                    Lưu
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {isDeliveryNoteEditModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-[#E4E3E0] border border-[#141414] p-8 w-full max-w-2xl space-y-6"
            >
              <h3 className="font-serif italic text-2xl">Chỉnh sửa dòng lệnh xuất</h3>
              <form onSubmit={saveDeliveryNoteItemEdit} className="space-y-4">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Mã hàng</label>
                    <input 
                      disabled
                      className="w-full bg-gray-200 border-b border-[#141414] py-1 text-sm outline-none opacity-60"
                      value={tempDeliveryNoteItem.item}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Tên hàng</label>
                    <input 
                      disabled
                      className="w-full bg-gray-200 border-b border-[#141414] py-1 text-sm outline-none opacity-60"
                      value={tempDeliveryNoteItem.materialName}
                    />
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">RPRO (OVN Production Order)</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={tempDeliveryNoteItem.ovnProductionOrder}
                      onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, ovnProductionOrder: e.target.value, lotNo: ''})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">NO (NoCode)</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={tempDeliveryNoteItem.noCode}
                      onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, noCode: e.target.value, lotNo: ''})}
                    />
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Qty ERP</label>
                    <input 
                      type="number"
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={tempDeliveryNoteItem.qtyErp}
                      onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, qtyErp: Number(e.target.value)})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Thực tế</label>
                    <input 
                      type="number"
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={tempDeliveryNoteItem.actualQty}
                      onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, actualQty: Number(e.target.value)})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Lot No</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={tempDeliveryNoteItem.lotNo}
                      onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, lotNo: e.target.value})}
                    />
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Remark</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={tempDeliveryNoteItem.remark}
                      onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, remark: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Vị trí</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={tempDeliveryNoteItem.location}
                      onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, location: e.target.value})}
                    />
                  </div>
                </div>
                {tempDeliveryNoteItem.stock === 'Không tìm thấy tồn kho phù hợp theo chỉ định' && (
                  <div className="flex items-center gap-2 text-red-600 bg-red-50 p-2 border border-red-200">
                    <AlertTriangle size={16} />
                    <span className="text-xs font-bold">Cảnh báo: Không tìm thấy tồn kho phù hợp theo chỉ định</span>
                  </div>
                )}
                <div className="flex gap-4 pt-4">
                  <button 
                    type="button"
                    onClick={() => setIsDeliveryNoteEditModalOpen(false)}
                    className="flex-1 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                  >
                    Hủy
                  </button>
                  <button 
                    type="submit"
                    className="flex-1 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                  >
                    Lưu thay đổi
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {isDeliveryNoteDeleteConfirmOpen && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-[#E4E3E0] border border-[#141414] p-8 w-full max-w-sm space-y-6 text-center"
            >
              <div className="w-16 h-16 bg-red-100 text-red-600 rounded-full flex items-center justify-center mx-auto">
                <Trash2 size={32} />
              </div>
              <div className="space-y-2">
                <h3 className="font-serif italic text-2xl">Xác nhận xóa dòng</h3>
                <p className="text-sm opacity-70">
                  Bạn có chắc chắn muốn xóa dòng này khỏi lệnh xuất kho không?
                </p>
              </div>
              <div className="flex gap-4 pt-2">
                <button 
                  onClick={() => {
                    setIsDeliveryNoteDeleteConfirmOpen(false);
                    setDeliveryNoteDeleteId(null);
                  }}
                  className="flex-1 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                >
                  Hủy
                </button>
                <button 
                  onClick={confirmDeleteDeliveryNoteItem}
                  className="flex-1 py-2 bg-red-600 text-white text-xs font-bold uppercase tracking-wider hover:bg-red-700 transition-colors"
                >
                  Xác nhận xóa
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {isDeleteConfirmOpen && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-[#E4E3E0] border border-[#141414] p-8 w-full max-w-sm space-y-6 text-center"
            >
              <div className="w-16 h-16 bg-red-100 text-red-600 rounded-full flex items-center justify-center mx-auto">
                <AlertTriangle size={32} />
              </div>
              <div className="space-y-2">
                <h3 className="font-serif italic text-2xl">Xác nhận xóa</h3>
                <p className="text-sm opacity-70">
                  Bạn có chắc chắn muốn xóa {deleteTarget?.id === 'wipe_all_data' ? 'TOÀN BỘ dữ liệu Tồn kho, Nhập kho, Xuất kho và Vị trí' : deleteTarget?.id === 'bulk' ? `hàng loạt (${selectedRows.length} mục)` : 'mục này'} không? Hành động này không thể hoàn tác.
                </p>
              </div>
              <div className="flex gap-4 pt-2">
                <button 
                  onClick={() => {
                    setIsDeleteConfirmOpen(false);
                    setDeleteTarget(null);
                  }}
                  className="flex-1 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                >
                  Hủy
                </button>
                <button 
                  onClick={confirmDelete}
                  className="flex-1 py-2 bg-red-600 text-white text-xs font-bold uppercase tracking-wider hover:bg-red-700 transition-colors"
                >
                  Xác nhận xóa
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {isSetupModalOpen && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white border border-[#141414] w-full max-w-2xl max-h-[90vh] overflow-y-auto"
          >
            <div className="bg-[#141414] text-[#E4E3E0] p-4 flex justify-between items-center sticky top-0">
              <h3 className="text-xs font-bold uppercase tracking-widest">HƯỚNG DẪN THIẾT LẬP SUPABASE</h3>
              <button onClick={() => setIsSetupModalOpen(false)} className="hover:opacity-70">
                <Plus size={20} className="rotate-45" />
              </button>
            </div>
            <div className="p-8 space-y-6">
              <section className="space-y-3">
                <h4 className="text-sm font-bold uppercase border-b border-gray-200 pb-2">Bước 1: Tạo dự án Supabase</h4>
                <p className="text-xs text-gray-600 leading-relaxed">
                  Truy cập <a href="https://supabase.com" target="_blank" rel="noreferrer" className="text-blue-600 underline">supabase.com</a>, tạo một dự án mới. Sau khi tạo xong, vào phần <strong>Project Settings &gt; API</strong> để lấy <code>Project URL</code> và <code>anon public key</code>.
                </p>
              </section>

              <section className="space-y-3">
                <h4 className="text-sm font-bold uppercase border-b border-gray-200 pb-2">Bước 2: Cấu hình biến môi trường</h4>
                <p className="text-xs text-gray-600 leading-relaxed">
                  Mở menu <strong>Settings</strong> (biểu tượng bánh răng) ở góc trên bên phải của AI Studio Build, sau đó thêm 2 biến sau:
                </p>
                <div className="bg-gray-100 p-3 font-mono text-[10px] space-y-1 border border-gray-200">
                  <p>VITE_SUPABASE_URL = [Project URL của bạn]</p>
                  <p>VITE_SUPABASE_ANON_KEY = [anon public key của bạn]</p>
                </div>
                <p className="text-[10px] text-amber-600 italic">
                  * Lưu ý: Nếu bạn deploy lên Vercel, hãy đảm bảo các biến này cũng được thêm vào Vercel Dashboard với tiền tố <code>VITE_</code>.
                </p>
              </section>

              <section className="space-y-3">
                <h4 className="text-sm font-bold uppercase border-b border-gray-200 pb-2">Bước 3: Khởi tạo Database</h4>
                <p className="text-xs text-gray-600 leading-relaxed">
                  Vào phần <strong>SQL Editor</strong> trong Supabase Dashboard, tạo một query mới và dán nội dung từ file <code>supabase_schema.sql</code> trong project này vào để tạo các bảng cần thiết.
                </p>
              </section>

              <div className="pt-4 flex justify-end">
                <button 
                  onClick={() => setIsSetupModalOpen(false)}
                  className="px-6 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-widest hover:opacity-90 transition-opacity"
                >
                  Đã hiểu
                </button>
              </div>
            </div>
          </motion.div>
        </div>
      )}
      <AnimatePresence>
        {notification && (
          <motion.div
            initial={{ opacity: 0, y: 50 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: 50 }}
            className={cn(
              "fixed bottom-8 right-8 px-6 py-4 rounded shadow-2xl flex items-center gap-3 z-50",
              notification.type === 'error' ? "bg-red-600 text-white" : "bg-green-600 text-white"
            )}
          >
            {notification.type === 'error' ? <AlertCircle size={20} /> : <CheckCircle2 size={20} />}
            <span className="font-bold">{notification.message}</span>
            <button onClick={() => setNotification(null)} className="ml-4 hover:opacity-70">
              <X size={16} />
            </button>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Hidden File Input */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileUpload}
        className="hidden"
        accept=".csv,.xlsx,.xls,.xlsm"
      />
        </div>
      </main>
    </div>
  );
}

function NavItem({ active, onClick, icon, label }: { active: boolean, onClick: () => void, icon: React.ReactNode, label: string }) {
  return (
    <button 
      onClick={onClick}
      className={cn(
        "w-full flex items-center gap-3 px-4 py-3 text-xs font-bold uppercase tracking-wider transition-all",
        active 
          ? "bg-[#141414] text-[#E4E3E0]" 
          : "hover:bg-[#141414]/5 opacity-60 hover:opacity-100"
      )}
    >
      {icon}
      {label}
    </button>
  );
}

function StatCard({ label, value, alert = false, bgColor = "bg-white/50", textColor = "text-[#141414]" }: { label: string, value: number | string, alert?: boolean, bgColor?: string, textColor?: string }) {
  return (
    <div className={cn(
      "border border-[#141414] p-4 h-32 flex flex-col gap-1",
      bgColor,
      alert && "border-red-500 bg-red-50/50"
    )}>
      <p className={cn("text-[9px] font-bold uppercase tracking-wider", textColor)}>{label}</p>
      <p className={cn(
        "text-lg font-bold",
        textColor,
        alert && "text-red-600"
      )}>
        {value}
      </p>
    </div>
  );
}

function StorageUsageBar({ showNotification }: { showNotification: (message: string, type?: 'success' | 'error') => void }) {
  const [usage, setUsage] = useState(0);
  const [isBackupModalOpen, setIsBackupModalOpen] = useState(false);
  const limit = 5 * 1024 * 1024; // 5MB limit for localStorage

  const calculateUsage = useCallback(() => {
    let total = 0;
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key) {
        total += (key.length + (localStorage.getItem(key) || '').length) * 2;
      }
    }
    setUsage(total);
  }, []);

  const percentage = Math.min(Math.round((usage / limit) * 100), 100);

  useEffect(() => {
    calculateUsage();
    const interval = setInterval(calculateUsage, 5000);
    window.addEventListener('storage', calculateUsage);
    
    return () => {
      clearInterval(interval);
      window.removeEventListener('storage', calculateUsage);
    };
  }, [calculateUsage]);

  const handleClearCache = () => {
    if (window.confirm('Bạn có chắc chắn muốn xóa bộ nhớ đệm? Dữ liệu sẽ được tải lại từ máy chủ.')) {
      api.clearCache();
      calculateUsage();
      window.location.reload();
    }
  };

  const handleBackupAndClear = () => {
    const backupData = api.getBackupData();
    const blob = new Blob([JSON.stringify(backupData, null, 2)], { type: 'application/json' });
    const timestamp = format(new Date(), 'yyyyMMdd_HHmmss');
    saveAs(blob, `warehouse_cache_backup_${timestamp}.json`);
    
    // Clear after a short delay to ensure download started
    setTimeout(() => {
      api.clearCache();
      calculateUsage();
      setIsBackupModalOpen(false);
      window.location.reload();
    }, 1000);
  };

  const handleFullSystemBackup = async () => {
    try {
      showNotification('Đang chuẩn bị dữ liệu sao lưu toàn hệ thống...', 'success');
      const fullData = await api.getFullBackupData();
      const blob = new Blob([JSON.stringify(fullData, null, 2)], { type: 'application/json' });
      const timestamp = format(new Date(), 'yyyyMMdd_HHmmss');
      saveAs(blob, `full_system_backup_${timestamp}.json`);
      showNotification('Đã tải về bản sao lưu toàn hệ thống thành công!', 'success');
    } catch (error) {
      console.error('Full backup error:', error);
      showNotification('Lỗi khi sao lưu toàn hệ thống.', 'error');
    }
  };

  const usageInMB = (usage / (1024 * 1024)).toFixed(2);

  return (
    <div className="space-y-2">
      <div className="flex justify-between items-end">
        <p className="text-[9px] font-bold uppercase tracking-widest opacity-50">Dung lượng lưu trữ</p>
        <p className="text-[10px] font-bold">{percentage}%</p>
      </div>
      <div className="h-1.5 w-full bg-gray-200 rounded-full overflow-hidden">
        <motion.div 
          initial={{ width: 0 }}
          animate={{ width: `${percentage}%` }}
          className={cn(
            "h-full transition-all duration-500",
            percentage > 90 ? "bg-red-500" : percentage > 70 ? "bg-amber-500" : "bg-green-500"
          )}
        />
      </div>
      <div className="flex justify-between items-center">
        <p className="text-[8px] opacity-40 italic">Đã dùng {usageInMB} MB / 5.00 MB</p>
        <div className="flex gap-2">
          <button 
            onClick={handleFullSystemBackup}
            className="text-[8px] font-bold uppercase tracking-tighter text-green-600 hover:underline flex items-center gap-0.5"
            title="Sao lưu toàn bộ dữ liệu từ máy chủ về máy tính"
          >
            <Download size={8} />
            Sao lưu hệ thống
          </button>
          <button 
            onClick={handleClearCache}
            className="text-[8px] font-bold uppercase tracking-tighter text-blue-600 hover:underline"
          >
            Xóa đệm
          </button>
        </div>
      </div>

      <AnimatePresence>
        {/* Modal removed as per user request */}
      </AnimatePresence>
    </div>
  );
}
