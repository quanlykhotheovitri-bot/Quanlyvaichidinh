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

  const [newProduct, setNewProduct] = useState<Partial<Product>>({
    sku: '', name: '', category: '', unit: '', minStock: 0, lotNo: '', ghiChu: '', designationCode: '', loaiChiDinh: ''
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
    const slashMatch = lot.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (slashMatch) {
      let [_, d, m, y] = slashMatch;
      const year = y.length === 2 ? '20' + y : y;
      const day = d.padStart(2, '0');
      const month = m.padStart(2, '0');
      return `${day}/${month}/${year}`;
    }
    
    const hyphenMatch = lot.match(/(\d{1,2})-(\d{1,2})-(\d{2,4})/);
    if (hyphenMatch) {
      let [_, d, m, y] = hyphenMatch;
      const year = y.length === 2 ? '20' + y : y;
      const day = d.padStart(2, '0');
      const month = m.padStart(2, '0');
      return `${day}/${month}/${year}`;
    }

    const isoMatch = lot.match(/(\d{2,4})-(\d{1,2})-(\d{1,2})/);
    if (isoMatch) {
      let [_, y, m, d] = isoMatch;
      const year = y.length === 2 ? '20' + y : y;
      const day = d.padStart(2, '0');
      const month = m.padStart(2, '0');
      return `${day}/${month}/${year}`;
    }
    return '';
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

    const timeoutId = setTimeout(async () => {
      if (!isOnline) {
        console.log('App is offline, skipping sync until network is back. Data is safely cached locally.');
        return;
      }

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
    setDeliveryNoteDeleteId(id);
    setIsDeliveryNoteDeleteConfirmOpen(true);
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

  const recalculateActualQty = (notes: DeliveryNoteItem[]) => {
    const groups = new Map<string, DeliveryNoteItem[]>();
    notes.forEach(item => {
      if (!groups.has(item.item)) groups.set(item.item, []);
      groups.get(item.item)!.push(item);
    });

    groups.forEach(items => {
      const totalQtyErp = items.reduce((sum, i) => sum + i.qtyErp, 0);
      const targetTotal = Math.ceil(totalQtyErp);
      const diff = targetTotal - totalQtyErp;
      
      let maxIdx = 0;
      let maxVal = -1;
      items.forEach((item, idx) => {
        if (item.qtyErp > maxVal) {
          maxVal = item.qtyErp;
          maxIdx = idx;
        }
      });

      items.forEach((item, idx) => {
        item.actualQty = (idx === maxIdx) ? Number((item.qtyErp + diff).toFixed(4)) : item.qtyErp;
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
      if (field === 'item' || field === 'lotNo') {
        updated = updated.map(item => {
          if (item.id === itemAtId) {
            return { ...item, location: getLocationByItemAndLot(item.item, item.lotNo) };
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
      if (item.assignedLots && item.assignedLots.length > 0) {
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
      } else if (item.actualQty && item.actualQty > 0) {
        const product = products.find(p => p.sku === item.item);
        if (product) {
          newTransactions.push({
            id: generateId(),
            productId: product.id,
            type: 'outbound',
            quantity: item.actualQty,
            date: today,
            partner: item.customerCode || item.noCode || 'Unknown',
            lotNo: item.lotNo,
            ghiChu: 'Xuất từ Phiếu giao nhận',
            designationCode: item.remark,
            loaiChiDinh: item.stock.includes('(') ? item.stock.split('(')[1].replace(')', '') : ''
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
    if (editingId) {
      updatedProduct = { ...newProduct, id: editingId } as Product;
      setProducts(products.map(p => p.id === editingId ? updatedProduct : p));
      setEditingId(null);
    } else {
      updatedProduct = {
        ...newProduct as Product,
        id: generateId(),
      };
      setProducts([...products, updatedProduct]);
    }
    
    try {
      await api.products.upsert(updatedProduct);
      showNotification('Sản phẩm đã được lưu.');
    } catch (error) {
      console.error('Error syncing product:', error);
      showNotification('Lỗi khi lưu sản phẩm.', 'error');
    }
    
    setIsProductModalOpen(false);
    setNewProduct({ sku: '', name: '', category: '', unit: '', minStock: 0, lotNo: '', ghiChu: '', designationCode: '', loaiChiDinh: '' });
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
        sku: normalizedRow['sku'] || row['SKU'] || row['sku'] || parsed?.sku || '',
        partner: normalizedRow['partner'] || normalizedRow['đối tác'] || row['Đối tác'] || row['partner'] || parsed?.partner || '',
        date: finalDate,
        location: normalizedRow['location'] || normalizedRow['vị trí'] || normalizedRow['vi tri'] || row['Vị trí'] || row['location'] || '',
        note: normalizedRow['note'] || normalizedRow['ghi chú'] || normalizedRow['cuộn'] || normalizedRow['cuon'] || row['Cuộn'] || row['Ghi chú'] || row['note'] || '',
        quantity: parseInt(String(normalizedRow['quantity'] || normalizedRow['số lượng'] || normalizedRow['so luong'] || row['Số lượng'] || row['quantity'] || '1')) || 1
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
    
    setLocationInventoryEntries(prev => {
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
              existing.sku = entry.sku || existing.sku;
              existing.partner = entry.partner || existing.partner;
              existing.date = entry.date || existing.date;
              existing.note = entry.note || existing.note;
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
        api.locationEntries.upsertAll(finalInventoryEntries),
      ];
      if (entriesToDeleteFromDb.length > 0) {
        entriesToDeleteFromDb.forEach(e => syncTasks.push(api.locationEntries.delete(e.id)));
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
      const productIds = idsToDelete.map(batchKey => {
        const item = inventory.find(i => i.id === batchKey);
        return item?.productId;
      }).filter(Boolean) as string[];
      const uniqueProductIds = [...new Set(productIds)];
      setProducts(prev => prev.filter(p => !uniqueProductIds.includes(p.id)));
      setTransactions(prev => prev.filter(t => !uniqueProductIds.includes(t.productId)));
      
      try {
        await Promise.all(uniqueProductIds.map(async (id) => {
          await api.transactions.deleteByProductId(id);
          await api.products.delete(id);
        }));
        showNotification(`Đã xóa ${uniqueProductIds.length} mặt hàng thành công.`);
      } catch (error: any) {
        console.error('Error in bulk delete:', error);
        showNotification('Lỗi khi xóa hàng loạt: ' + (error.message || ''), 'error');
        loadData();
      }
    } else if (currentTab === 'inbound' || currentTab === 'outbound') {
      setTransactions(prev => prev.filter(t => !idsToDelete.includes(t.id)));
      try {
        await Promise.all(idsToDelete.map(id => api.transactions.delete(id)));
        showNotification('Đã xóa các giao dịch thành công.');
      } catch (error) {
        showNotification('Lỗi khi xóa giao dịch.', 'error');
        loadData();
      }
    } else if (currentTab === 'customers') {
      setCustomers(prev => prev.filter(c => !idsToDelete.includes(c.id)));
      try {
        await Promise.all(idsToDelete.map(id => api.customers.delete(id)));
        showNotification('Đã xóa các khách hàng thành công.');
      } catch (error) {
        showNotification('Lỗi khi xóa khách hàng.', 'error');
        loadData();
      }
    } else if (currentTab === 'location') {
      if (subTab === 'input' || subTab === 'output') {
        setLocationEntries(prev => prev.filter(e => !idsToDelete.includes(e.id)));
        try {
          await Promise.all(idsToDelete.map(id => api.locationEntries.delete(id)));
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
          await Promise.all(uniqueQrcodes.map(qrcode => api.locationEntries.deleteByQRCode(qrcode)));
          showNotification('Đã xóa các mục tồn vị trí thành công.');
        } catch (error) {
          showNotification('Lỗi khi xóa tồn vị trí.', 'error');
          loadData();
        }
      }
    } else if (currentTab === 'deliveryNote') {
      if (deliveryNoteSubTab === 'preview') {
        setDeliveryNotes(prev => prev.filter(item => !idsToDelete.includes(item.id)));
        try {
          await Promise.all(idsToDelete.map(id => api.deliveryNotes.delete(id)));
          showNotification('Đã xóa các mục phiếu giao nhận thành công.');
        } catch (error) {
          console.error('Error in bulk delete delivery notes:', error);
          showNotification('Lỗi khi xóa phiếu giao nhận.', 'error');
          loadData();
        }
      } else {
        setSavedDeliveryNotes(prev => prev.filter(n => !idsToDelete.includes(n.id)));
        try {
          await Promise.all(idsToDelete.map(id => api.savedDeliveryNotes.delete(id)));
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

  const handleSaveAll = async () => {
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
      showNotification('Dữ liệu đã được lưu trữ thành công!');
    } catch (error) {
      console.error('Lỗi khi lưu dữ liệu:', error);
      showNotification('Có lỗi xảy ra khi lưu dữ liệu. Vui lòng kiểm tra lại kết nối.', 'error');
    } finally {
      setIsSaving(false);
    }
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
    if (deleteTarget.id === 'bulk') {
      await handleBulkDelete();
    } else if (deleteTarget.id === 'wipe_all_data') {
      await handleWipeAllData();
    } else {
      await handleDelete(deleteTarget.id, deleteTarget.type as any);
    }
    setIsDeleteConfirmOpen(false);
    setDeleteTarget(null);
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
      const type = item.loaiChiDinh || 'N/A';
      summary.set(type, (summary.get(type) || 0) + item.currentStock);
    });
    
    return Array.from(summary.entries())
      .map(([type, qty]) => `${type}: ${qty.toLocaleString()}`)
      .join(' | ');
  }, [inventory]);

  const getLocationByItemAndLot = useCallback((sku: string, lotNo: string) => {
    if (!sku) return 'Chưa có vị trí';
    
    const skuLower = sku.toLowerCase().trim();
    const normalizedLotDate = normalizeDateForMatching(lotNo || '');
    
    const matches = locationInventoryEntries.filter(e => 
      (e.sku || '').toLowerCase().trim() === skuLower && 
      (e.date || '') === normalizedLotDate
    );
    
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

    const workingInventory = inventory.map(item => ({
      ...item,
      tempStock: item.currentStock
    }));

    const updatedNotes = deliveryNotes.map(item => {
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

      let remainingNeeded = item.actualQty || item.qtyErp;
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
      <input 
        type="file" 
        ref={fileInputRef} 
        onChange={handleFileUpload} 
        className="hidden" 
        accept=".csv,.xlsx,.xls,.xlsm"
      />
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
        "fixed inset-y-0 left-0 z-50 w-64 border-r border-[#141414] bg-[#E4E3E0] flex flex-col transition-transform duration-300 lg:relative lg:translate-x-0",
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

        <div className="p-6 mt-auto border-t border-[#141414]">
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
        <header className="h-20 border-b border-[#141414] flex items-center justify-between px-4 sm:px-8 bg-[#E4E3E0]/80 backdrop-blur-sm z-10 gap-4">
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
              <button 
                onClick={handleSaveAll}
                disabled={isSaving}
                className={`flex items-center gap-2 px-3 sm:px-4 py-2 text-[10px] font-bold uppercase tracking-wider transition-all disabled:opacity-80 ${!isOnline ? 'bg-red-600 text-white' : isSaving ? 'bg-[#ff9900] text-black' : 'bg-[#141414] text-[#E4E3E0] hover:opacity-90'}`}
                title={!isOnline ? "Hoạt động Ngoại Tuyến (Offline)" : "Hệ thống đã bật tự động lưu"}
              >
                {!isOnline ? (
                  <WifiOff size={14} />
                ) : isSaving ? (
                  <div className="w-3 h-3 border-2 border-black border-t-transparent rounded-full animate-spin" />
                ) : (
                  <Save size={14} />
                )}
                <span className="hidden md:inline">{isSaving ? 'Đang lưu...' : 'Lưu tất cả'}</span>
              </button>
              
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
                <span className="hidden md:inline">Cập nhật file</span>
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
                        setNewTransaction(prev => ({ ...prev, type: activeTab as 'inbound' | 'outbound' }));
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
                    <button 
                      onClick={handleExportInventory}
                      className="flex items-center gap-2 px-4 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                    >
                      <Download size={14} />
                      <span className="hidden sm:inline">Xuất Excel</span>
                    </button>
                    <button 
                      onClick={() => {
                        setEditingId(null);
                        setNewProduct({ sku: '', name: '', category: '', unit: '', minStock: 0, lotNo: '', ghiChu: '', designationCode: '', loaiChiDinh: '' });
                        setIsProductModalOpen(true);
                      }}
                      className="flex items-center gap-2 px-4 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <Plus size={14} />
                      <span className="hidden sm:inline">Thêm mặt hàng</span>
                    </button>
                  </div>
                </div>

                <div className="border border-[#141414] overflow-x-auto max-h-[70vh]">
                  <table className="w-full border-collapse min-w-[1000px]">
                    <thead className="sticky top-0 z-10 shadow-sm">
                      <tr className="bg-[#001F3F] text-white text-[11px] uppercase tracking-wider">
                        <th className="border border-[#141414] p-3 w-10">
                          <button onClick={() => {
                            const currentIds = filteredInventory.map(item => item.id);
                            if (selectedRows.length === currentIds.length) setSelectedRows([]);
                            else setSelectedRows(currentIds);
                          }}>
                            {selectedRows.length > 0 ? <CheckSquare size={14} /> : <Square size={14} />}
                          </button>
                        </th>
                        <th className="border border-[#141414] p-3 text-left">Mã Hàng</th>
                        <th className="border border-[#141414] p-3 text-left">Tên hàng</th>
                        <th className="border border-[#141414] p-3 text-left">Lot no</th>
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
                            <td className="border border-[#141414] p-3 text-center font-bold">{item.currentStock.toLocaleString()}</td>
                            <td className="border border-[#141414] p-3 opacity-60">{item.loaiChiDinh}</td>
                            <td className="border border-[#141414] p-3">{item.ghiChu}</td>
                            <td className="border border-[#141414] p-3 font-mono">{item.designationCode}</td>
                            <td className="border border-[#141414] p-3 text-center">
                              <div className="flex items-center justify-center gap-2">
                                <button 
                                  onClick={() => {
                                    setEditingId(item.productId);
                                    setNewProduct(products.find(p => p.id === item.productId) || {});
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
                        <th className="border border-[#141414] p-3 text-left">Mã KH</th>
                        <th className="border border-[#141414] p-3 text-left">Tên khách hàng</th>
                        <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                          Thao tác
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredCustomers.map(c => {
                        const isSelected = selectedRows.includes(c.id);
                        return (
                          <tr key={c.id} className={cn(
                            "bg-white text-xs hover:bg-gray-50 transition-colors",
                            isSelected && "bg-blue-50"
                          )}>
                            <td className="border border-[#141414] p-3 text-center">
                              <button onClick={() => toggleRowSelection(c.id)}>
                                {isSelected ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} />}
                              </button>
                            </td>
                            <td className="border border-[#141414] p-3 font-mono">{c.code}</td>
                            <td className="border border-[#141414] p-3 font-bold">{c.name}</td>
                            <td className="border border-[#141414] p-3 text-center">
                              <div className="flex items-center justify-center gap-2">
                                <button 
                                  onClick={() => {
                                    setEditingId(c.id);
                                    setNewCustomer(c);
                                    setIsCustomerModalOpen(true);
                                  }}
                                  className="p-1 hover:bg-gray-200 rounded transition-colors"
                                >
                                  <Edit2 size={14} className="text-blue-600" />
                                </button>
                                <button 
                                  onClick={() => {
                                    setDeleteTarget({ id: c.id, type: 'customer' });
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
                <div className="flex border-b border-[#141414] mb-6 no-print">
                  <button 
                    onClick={() => setDeliveryNoteSubTab('preview')}
                    className={cn(
                      "px-8 py-4 text-xs font-bold uppercase tracking-widest transition-all",
                      deliveryNoteSubTab === 'preview' ? "bg-[#141414] text-[#E4E3E0]" : "hover:bg-[#141414]/5"
                    )}
                  >
                    SOẠN PHIẾU GIAO HÀNG
                  </button>
                  <button 
                    onClick={() => setDeliveryNoteSubTab('history')}
                    className={cn(
                      "px-8 py-4 text-xs font-bold uppercase tracking-widest transition-all",
                      deliveryNoteSubTab === 'history' ? "bg-[#141414] text-[#E4E3E0]" : "hover:bg-[#141414]/5"
                    )}
                  >
                    LỊCH SỬ PHIẾU
                  </button>
                </div>

                {deliveryNoteSubTab === 'preview' ? (
                  <>
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 no-print mb-6">
                      <div className="flex flex-wrap gap-2 sm:gap-3 w-full sm:w-auto">
                        <button 
                          onClick={() => fileInputRef.current?.click()}
                          className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                        >
                          <FileUp size={14} />
                          TẢI FILE LỆNH (ERP)
                        </button>
                        <button 
                          onClick={autoAssignLotsFromUI}
                          className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2 bg-blue-600 text-white text-xs font-bold uppercase tracking-wider hover:bg-blue-700 transition-colors"
                        >
                          <CheckCircle2 size={14} />
                          GẮN LOT TỰ ĐỘNG
                        </button>
                        <button 
                          onClick={handlePostDeliveryNote}
                          className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2 bg-green-600 text-white text-xs font-bold uppercase tracking-wider hover:bg-green-700 transition-colors"
                        >
                          <ArrowUpFromLine size={14} />
                          POST LỆNH XUẤT KHO
                        </button>
                        <button 
                          onClick={() => window.print()}
                          className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-4 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                        >
                          <Printer size={14} />
                          IN PHIẾU
                        </button>
                        {selectedRows.length > 0 && (
                          <button
                            onClick={() => {
                              setDeleteTarget({ id: 'bulk', type: 'deliveryNote' });
                              setIsDeleteConfirmOpen(true);
                            }}
                            className="flex items-center gap-2 px-4 py-2 bg-red-600 text-white text-xs font-bold uppercase tracking-wider hover:bg-red-700 transition-colors"
                          >
                            <Trash2 size={14} />
                            Xóa ({selectedRows.length})
                          </button>
                        )}
                      </div>
                    </div>

                    <div id="print-area" className="bg-white p-8 sm:p-12 border border-[#141414] shadow-sm max-w-5xl mx-auto text-[#141414]">
                      {/* Header Logo - Only Text for now but styled */}
                      <div className="flex justify-between items-start mb-10">
                        <div className="space-y-1">
                          <h2 className="text-xl font-bold tracking-tighter">ORTHOLITE VIETNAM</h2>
                          <p className="text-[10px] uppercase font-bold text-gray-400">Inventory Management System</p>
                        </div>
                        <div className="text-right space-y-1">
                          <p className="text-[11px] font-bold uppercase tracking-wider">Mã tài liệu: {deliveryNoteHeader.documentCode}</p>
                          <p className="text-[10px] italic text-gray-500">Ngày in: {format(new Date(), 'dd/MM/yyyy HH:mm')}</p>
                        </div>
                      </div>

                      <div className="text-center space-y-2 mb-12">
                        <h1 className="text-3xl font-serif italic font-bold uppercase tracking-tight">PHIẾU GIAO NHẬN NỘI BỘ</h1>
                        <p className="text-sm font-bold opacity-60">(INTERNAL DELIVERY NOTE)</p>
                      </div>

                      <div className="grid grid-cols-2 gap-x-12 gap-y-6 mb-12 text-sm">
                        <div className="flex border-b border-gray-200 pb-2 flex-col">
                          <span className="text-[10px] font-bold uppercase tracking-wider opacity-40 mb-1">BỘ PHẬN GIAO (DEPT):</span>
                          <input 
                            type="text" 
                            value={deliveryNoteHeader.dept}
                            onChange={(e) => setDeliveryNoteHeader(prev => ({ ...prev, dept: e.target.value }))}
                            className="bg-transparent border-none outline-none font-bold placeholder:opacity-20"
                            placeholder="Nhập bộ phận..."
                          />
                        </div>
                        <div className="flex border-b border-gray-200 pb-2 flex-col">
                          <span className="text-[10px] font-bold uppercase tracking-wider opacity-40 mb-1">NGÀY (DATE):</span>
                          <input 
                            type="text" 
                            value={deliveryNoteHeader.date}
                            onChange={(e) => setDeliveryNoteHeader(prev => ({ ...prev, date: e.target.value }))}
                            className="bg-transparent border-none outline-none font-bold placeholder:opacity-20"
                            placeholder="dd/mm/yyyy"
                          />
                        </div>
                        <div className="flex border-b border-gray-200 pb-2 flex-col">
                          <span className="text-[10px] font-bold uppercase tracking-wider opacity-40 mb-1">BỘ PHẬN NHẬN (TO):</span>
                          <input 
                            type="text" 
                            value={deliveryNoteHeader.to}
                            onChange={(e) => setDeliveryNoteHeader(prev => ({ ...prev, to: e.target.value }))}
                            className="bg-transparent border-none outline-none font-bold placeholder:opacity-20"
                            placeholder="Nhập người nhận/bộ phận..."
                          />
                        </div>
                      </div>

                      <div className="overflow-x-auto mb-12">
                        <table className="w-full border-collapse">
                          <thead>
                            <tr className="bg-gray-50 text-[10px] uppercase font-bold">
                              <th className="border-2 border-[#141414] p-2 no-print w-8 h-8">
                                <button onClick={() => {
                                  const currentIds = filteredDeliveryNotes.map(item => item.id);
                                  if (selectedRows.length === currentIds.length) setSelectedRows([]);
                                  else setSelectedRows(currentIds);
                                }}>
                                  {selectedRows.length > 0 ? <CheckSquare size={14} /> : <Square size={14} />}
                                </button>
                              </th>
                              <th className="border-2 border-[#141414] p-2 w-10">STT</th>
                              <th className="border-2 border-[#141414] p-2 text-left">Mô tả hàng hóa</th>
                              <th className="border-2 border-[#141414] p-2">Mã HH</th>
                              <th className="border-2 border-[#141414] p-2 w-16">ĐVT</th>
                              <th className="border-2 border-[#141414] p-2 w-20">S.Lượng</th>
                              <th className="border-2 border-[#141414] p-2 w-24">Lot No.</th>
                              <th className="border-2 border-[#141414] p-2 w-32">Vị Trí</th>
                              <th className="border-2 border-[#141414] p-2 text-left">Ghi chú</th>
                              <th className="border-2 border-[#141414] p-2 no-print w-20">Xử lý</th>
                            </tr>
                          </thead>
                          <tbody className="text-xs">
                            {filteredDeliveryNotes.map((item, index) => {
                              const isSelected = selectedRows.includes(item.id);
                              return (
                                <tr key={item.id} className={cn(
                                  "hover:bg-gray-50 group",
                                  isSelected && "bg-blue-50"
                                )}>
                                  <td className="border-2 border-[#141414] p-2 text-center no-print">
                                    <button onClick={() => toggleRowSelection(item.id)}>
                                      {isSelected ? <CheckSquare size={14} /> : <Square size={14} />}
                                    </button>
                                  </td>
                                  <td className="border-2 border-[#141414] p-2 text-center font-bold">{index + 1}</td>
                                  <td className="border-2 border-[#141414] p-2">
                                    <p className="font-bold">{item.materialName}</p>
                                    <p className="text-[9px] opacity-40 uppercase">Brand: {item.brand} | Cust: {item.customerCode}</p>
                                  </td>
                                  <td className="border-2 border-[#141414] p-2 font-mono text-center text-[10px]">{item.item}</td>
                                  <td className="border-2 border-[#141414] p-2 text-center">{item.unit}</td>
                                  <td className="border-2 border-[#141414] p-2 text-center font-bold italic">{item.actualQty.toLocaleString()}</td>
                                  <td className="border-2 border-[#141414] p-2">
                                    {item.assignedLots && item.assignedLots.length > 0 ? (
                                      <div className="space-y-1">
                                        {item.assignedLots.map((lot, idx) => (
                                          <div key={idx} className="flex justify-between items-center gap-2 border-b border-gray-100 last:border-0 pb-1">
                                            <span className="font-mono">{lot.lotNo}</span>
                                            <span className="font-bold opacity-60">({lot.qty.toLocaleString()})</span>
                                          </div>
                                        ))}
                                      </div>
                                    ) : (
                                      <div className="text-red-500 font-bold text-[9px] text-center">CHƯA GẮN LOT</div>
                                    )}
                                  </td>
                                  <td className="border-2 border-[#141414] p-2">
                                    {item.assignedLots && item.assignedLots.length > 0 ? (
                                      <div className="text-center font-bold">
                                        {[...new Set(item.assignedLots.map(l => l.location))].filter(Boolean).join(', ') || 'N/A'}
                                      </div>
                                    ) : (
                                      <div className="text-center text-gray-300">N/A</div>
                                    )}
                                  </td>
                                  <td className="border-2 border-[#141414] p-2">
                                    {item.assignedLots && item.assignedLots.length > 0 ? (
                                      <div className="space-y-1">
                                        {item.assignedLots.map((lot, idx) => (
                                          <p key={idx} className="text-[9px] opacity-60 leading-tight">
                                            {lot.stock}
                                          </p>
                                        ))}
                                      </div>
                                    ) : (
                                      <p className="text-red-400 italic text-[9px]">{item.stock}</p>
                                    )}
                                  </td>
                                  <td className="border-2 border-[#141414] p-2 no-print">
                                    <div className="flex items-center justify-center gap-1 opacity-10 group-hover:opacity-100 transition-opacity">
                                      <button 
                                        onClick={() => handleEditDeliveryNoteItemClick(item.id)}
                                        className="p-1 hover:bg-gray-200 rounded text-blue-600"
                                      >
                                        <Edit2 size={12} />
                                      </button>
                                      <button 
                                        onClick={() => handleDeleteDeliveryNoteItem(item.id)}
                                        className="p-1 hover:bg-gray-200 rounded text-red-600"
                                      >
                                        <Trash2 size={12} />
                                      </button>
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                            
                            {/* Filling empty rows to maintain form length if needed */}
                            {filteredDeliveryNotes.length < 10 && [...Array(Math.max(0, 10 - filteredDeliveryNotes.length))].map((_, i) => (
                              <tr key={`empty-${i}`} className="h-10">
                                <td className="border-2 border-[#141414] p-2 no-print"></td>
                                <td className="border-2 border-[#141414] p-2"></td>
                                <td className="border-2 border-[#141414] p-2"></td>
                                <td className="border-2 border-[#141414] p-2"></td>
                                <td className="border-2 border-[#141414] p-2"></td>
                                <td className="border-2 border-[#141414] p-2"></td>
                                <td className="border-2 border-[#141414] p-2"></td>
                                <td className="border-2 border-[#141414] p-2"></td>
                                <td className="border-2 border-[#141414] p-2"></td>
                                <td className="border-2 border-[#141414] p-2 no-print"></td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>

                      <div className="grid grid-cols-3 gap-8 text-center mt-16 font-bold uppercase tracking-widest text-[10px]">
                        <div className="space-y-24">
                          <p>NGƯỜI GIAO (ISSUER)</p>
                          <div className="flex flex-col items-center">
                            <p className="border-t border-[#141414] w-32 pt-2 opacity-30 font-normal normal-case italic">Ký & ghi rõ họ tên</p>
                          </div>
                        </div>
                        <div className="space-y-24">
                          <p>THỦ KHO (KEEPER)</p>
                          <div className="flex flex-col items-center">
                            <p className="border-t border-[#141414] w-32 pt-2 opacity-30 font-normal normal-case italic">Ký & ghi rõ họ tên</p>
                          </div>
                        </div>
                        <div className="space-y-24">
                          <p>NGƯỜI NHẬN (RECEIVER)</p>
                          <div className="flex flex-col items-center">
                            <p className="border-t border-[#141414] w-32 pt-2 opacity-30 font-normal normal-case italic">Ký & ghi rõ họ tên</p>
                          </div>
                        </div>
                      </div>
                    </div>
                  </>
                ) : (
                  <div className="space-y-4">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4">
                      <h2 className="font-serif italic text-xl">Lịch sử phiếu đã lưu</h2>
                      <div className="flex items-center gap-3">
                        <div className="relative">
                          <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                          <input
                            type="text"
                            placeholder="Tìm kiếm phiếu..."
                            value={searchQuery}
                            onChange={(e) => setSearchQuery(e.target.value)}
                            className="pl-10 pr-4 py-2 border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-64"
                          />
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                      {savedDeliveryNotes.map(note => (
                        <div key={note.id} className="border-2 border-[#141414] p-6 bg-white hover:shadow-lg transition-all relative group">
                          <div className="flex justify-between items-start mb-4">
                            <div>
                              <div className="flex items-center gap-2 mb-1">
                                <FileText size={16} className="text-[#1a5f7a]" />
                                <h3 className="font-bold text-sm">PHIẾU : {note.date}</h3>
                              </div>
                              <p className="text-[10px] font-bold uppercase tracking-widest opacity-40">Mã ID: {note.id.slice(0, 8)}...</p>
                            </div>
                            <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                              <button 
                                onClick={() => {
                                  setDeliveryNotes(note.items);
                                  setDeliveryNoteSubTab('preview');
                                }}
                                className="p-2 hover:bg-[#141414] hover:text-[#E4E3E0] rounded transition-all"
                                title="Mở lại phiếu"
                              >
                                <Plus size={16} />
                              </button>
                              <button 
                                onClick={() => handleDelete(note.id, 'savedDeliveryNote')}
                                className="p-2 hover:bg-red-600 hover:text-white rounded transition-all"
                                title="Xóa phiếu"
                              >
                                <Trash2 size={16} />
                              </button>
                            </div>
                          </div>
                          <div className="space-y-2 border-t border-gray-100 pt-4">
                            <p className="text-xs font-bold">Tổng số mặt hàng: {note.items.length}</p>
                            <p className="text-xs font-bold text-[#1a5f7a]">
                              Tổng số lượng: {note.items.reduce((sum, i) => sum + (i.actualQty || 0), 0).toLocaleString()}
                            </p>
                          </div>
                        </div>
                      ))}
                      {savedDeliveryNotes.length === 0 && (
                        <div className="col-span-full py-20 text-center border-2 border-dashed border-[#141414] opacity-30">
                          <FileText size={48} className="mx-auto mb-4" />
                          <p className="text-xs font-bold uppercase tracking-widest">Chưa có lịch sử phiếu được lưu</p>
                        </div>
                      )}
                    </div>
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
                <div className="flex border-b border-[#141414] mb-6 no-print">
                  <button 
                    onClick={() => setLocationSubTab('input')}
                    className={cn(
                      "px-8 py-4 text-xs font-bold uppercase tracking-widest transition-all",
                      locationSubTab === 'input' ? "bg-blue-600 text-white" : "hover:bg-blue-50 text-blue-600/60"
                    )}
                  >
                    LỊCH SỬ NHẬP VỊ TRÍ
                  </button>
                  <button 
                    onClick={() => setLocationSubTab('output')}
                    className={cn(
                      "px-8 py-4 text-xs font-bold uppercase tracking-widest transition-all",
                      locationSubTab === 'output' ? "bg-orange-600 text-white" : "hover:bg-orange-50 text-orange-600/60"
                    )}
                  >
                    LỊCH SỬ XUẤT VỊ TRÍ
                  </button>
                  <button 
                    onClick={() => setLocationSubTab('inventory')}
                    className={cn(
                      "px-8 py-4 text-xs font-bold uppercase tracking-widest transition-all",
                      locationSubTab === 'inventory' ? "bg-green-600 text-white" : "hover:bg-green-50 text-green-600/60"
                    )}
                  >
                    BÁO CÁO TỒN VỊ TRÍ
                  </button>
                </div>

                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <h2 className="font-serif italic text-2xl capitalize">
                    {locationSubTab === 'inventory' ? 'QUẢN LÝ VỊ TRÍ (STOCK)' : 'NHẬP/XUẤT NHANH (LOG)'}
                  </h2>
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
                          Chế độ NHẬP
                        </button>
                        <button
                          onClick={() => setScanMode('OUTPUT')}
                          className={`flex-1 sm:flex-none px-6 py-2 text-[10px] font-bold uppercase tracking-widest transition-all border border-[#141414] ${
                            scanMode === 'OUTPUT' 
                              ? 'bg-orange-600 text-white opacity-100' 
                              : 'bg-white text-orange-600 opacity-30 hover:opacity-50'
                          }`}
                        >
                          Chế độ XUẤT
                        </button>
                      </div>
                    </div>
                    <div className="space-y-2">
                      <p className="text-[10px] text-gray-500 italic">Mẹo: Quét mã vị trí (bắt đầu bằng FB) trước để định vị, sau đó quét các mã cuộn vải.</p>
                      <div className="flex gap-4">
                        <div className="flex-1">
                          <textarea
                            rows={4}
                            placeholder="Dán mã QR hoặc quét tại đây... (Ví dụ: FB01-A-01 rồi mã cuộn)"
                            className="w-full p-4 border border-[#141414] text-sm font-mono focus:outline-none focus:ring-1 focus:ring-[#141414] bg-gray-50"
                            value={scanInput}
                            onChange={(e) => setScanInput(e.target.value)}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter' && !e.shiftKey) {
                                e.preventDefault();
                                handleProcessScanInput();
                              }
                            }}
                          />
                        </div>
                        <div className="w-48 flex flex-col gap-2">
                          <div className="flex-1 bg-gray-100 border border-[#141414] p-3 flex flex-col justify-center items-center text-center">
                            <span className="text-[10px] font-bold uppercase tracking-wider opacity-40">Vị trí hiện tại</span>
                            <span className="text-xl font-bold font-mono">{currentLocation || 'N/A'}</span>
                          </div>
                          <button
                            onClick={handleProcessScanInput}
                            className={cn(
                              "py-3 text-white text-xs font-bold uppercase tracking-widest transition-all",
                              scanMode === 'INPUT' ? "bg-blue-600 hover:bg-blue-700" : "bg-orange-600 hover:bg-orange-700"
                            )}
                          >
                            Xử lý
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                <div className="border border-[#141414] overflow-x-auto max-h-[70vh]">
                  <table className="w-full border-collapse min-w-[1000px]">
                    <thead className="sticky top-0 z-10 shadow-sm">
                      <tr className={cn(
                        "text-white text-[11px] uppercase tracking-wider",
                        locationSubTab === 'input' ? "bg-blue-600" : (locationSubTab === 'output' ? "bg-orange-600" : "bg-green-600")
                      )}>
                        <th className="border border-[#141414] p-3 w-10">
                          <button onClick={() => {
                            const currentIds = (locationSubTab === 'inventory' ? filteredLocationInventoryEntries : filteredLocationEntries).map(e => e.id);
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
                        <th className="border border-[#141414] p-3 text-center bg-yellow-400 text-[#141414]">Vị trí</th>
                        <th className="border border-[#141414] p-3 text-left">Cuộn/Ghi chú</th>
                        {locationSubTab === 'inventory' && <th className="border border-[#141414] p-3 text-center">Số lượng</th>}
                        <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                          Thao tác
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {(locationSubTab === 'inventory' ? filteredLocationInventoryEntries : filteredLocationEntries).map(entry => {
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
                            <td className="border border-[#141414] p-3 font-mono">{entry.date}</td>
                            <td className="border border-[#141414] p-3 text-center font-bold bg-yellow-50">{entry.location}</td>
                            <td className="border border-[#141414] p-3 italic">{entry.note}</td>
                            {locationSubTab === 'inventory' && <td className="border border-[#141414] p-3 text-center font-bold text-red-600">{entry.quantity}</td>}
                            <td className="border border-[#141414] p-3 text-center">
                              <div className="flex items-center justify-center gap-2">
                                <button 
                                  onClick={() => {
                                    setEditingId(entry.id);
                                    setNewLocationEntry(entry);
                                    setIsLocationModalOpen(true);
                                  }}
                                  className="p-1 hover:bg-gray-200 rounded transition-colors"
                                >
                                  <Edit2 size={14} className="text-blue-600" />
                                </button>
                                <button 
                                  onClick={() => {
                                    setDeleteTarget({ 
                                      id: entry.id, 
                                      type: 'location', 
                                      qrcode: locationSubTab === 'inventory' ? entry.qrcode : undefined 
                                    });
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
          </AnimatePresence>
        </div>
      </main>

      {/* Global Notification */}
      <AnimatePresence>
        {notification && (
          <motion.div 
            initial={{ opacity: 0, y: 50 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: 50 }}
            className="fixed bottom-8 left-1/2 -translate-x-1/2 z-[100]"
          >
            <div className={cn(
              "px-6 py-4 border-2 border-[#141414] font-bold text-xs uppercase tracking-widest flex items-center gap-3 shadow-2xl",
              notification.type === 'success' ? "bg-white text-green-600" : "bg-red-600 text-white"
            )}>
              {notification.type === 'success' ? <CheckCircle2 size={18} /> : <AlertCircle size={18} />}
              {notification.message}
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Settings/Setup Modal */}
      <Modal 
        isOpen={isSetupModalOpen} 
        onClose={() => setIsSetupModalOpen(false)}
        title="THIẾT LẬP HỆ THỐNG (SUPABASE)"
      >
        <div className="space-y-6">
          <div className="p-4 bg-gray-100 border-l-4 border-amber-500 text-xs">
            <p className="font-bold mb-2 uppercase">Lưu ý quan trọng:</p>
            <p>Hệ thống sử dụng Supabase để lưu trữ dữ liệu an toàn. Nếu bạn chưa cấu hình, dữ liệu sẽ chỉ được lưu tạm thời trên trình duyệt của bạn.</p>
          </div>
          <div className="space-y-4">
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Supabase URL</label>
              <input 
                type="text" 
                value={process.env.VITE_SUPABASE_URL || ''}
                readOnly
                className="w-full p-3 bg-gray-50 border border-[#141414] text-xs font-mono opacity-50"
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Supabase Anon Key</label>
              <input 
                type="text" 
                value={process.env.VITE_SUPABASE_ANON_KEY ? '••••••••••••••••' : ''}
                readOnly
                className="w-full p-3 bg-gray-50 border border-[#141414] text-xs font-mono opacity-50"
              />
            </div>
          </div>
          <div className="pt-4 border-t border-gray-100">
            <p className="text-[10px] text-gray-500 leading-relaxed">
              * Để thay đổi cấu hình, vui lòng cập nhật file <code>.env</code> trong thư mục gốc của dự án và khởi động lại ứng dụng.
            </p>
          </div>
          <div className="flex justify-end gap-3 pt-6">
            <button 
              onClick={() => setIsSetupModalOpen(false)}
              className="px-6 py-3 border border-[#141414] text-xs font-bold uppercase tracking-widest hover:bg-[#141414] hover:text-[#E4E3E0] transition-all"
            >
              Đóng
            </button>
          </div>
        </div>
      </Modal>

      {/* Product Modal */}
      <Modal 
        isOpen={isProductModalOpen} 
        onClose={() => setIsProductModalOpen(false)}
        title={editingId ? "CHỈNH SỬA MẶT HÀNG" : "THÊM MẶT HÀNG MỚI"}
      >
        <form onSubmit={handleAddProduct} className="space-y-6">
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Mã Hàng (SKU)</label>
              <input 
                type="text" required 
                value={newProduct.sku}
                onChange={e => setNewProduct({...newProduct, sku: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Tên Sản Phẩm</label>
              <input 
                type="text" required 
                value={newProduct.name}
                onChange={e => setNewProduct({...newProduct, name: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Loại Hàng</label>
              <input 
                type="text" 
                value={newProduct.category}
                onChange={e => setNewProduct({...newProduct, category: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Đơn Vị Tính</label>
              <input 
                type="text" 
                value={newProduct.unit}
                onChange={e => setNewProduct({...newProduct, unit: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Min Stock</label>
              <input 
                type="number" 
                value={newProduct.minStock}
                onChange={e => setNewProduct({...newProduct, minStock: Number(e.target.value)})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Loại chỉ định</label>
              <input 
                type="text" 
                value={newProduct.loaiChiDinh}
                onChange={e => setNewProduct({...newProduct, loaiChiDinh: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Mã chỉ định</label>
              <input 
                type="text" 
                value={newProduct.designationCode}
                onChange={e => setNewProduct({...newProduct, designationCode: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Lot No</label>
              <input 
                type="text" 
                value={newProduct.lotNo}
                onChange={e => setNewProduct({...newProduct, lotNo: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
          </div>
          <div>
            <label className="block text-[10px] font-bold uppercase mb-2">Ghi Chú</label>
            <textarea 
              value={newProduct.ghiChu}
              onChange={e => setNewProduct({...newProduct, ghiChu: e.target.value})}
              className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              rows={2}
            />
          </div>
          <div className="flex justify-end gap-3 pt-6">
            <button 
              type="button" 
              onClick={() => setIsProductModalOpen(false)}
              className="px-6 py-3 border border-[#141414] text-xs font-bold uppercase tracking-widest hover:bg-gray-100 transition-all"
            >
              Hủy
            </button>
            <button 
              type="submit"
              className="px-8 py-3 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-widest hover:opacity-90 transition-all"
            >
              {editingId ? "Cập nhật" : "Lưu mặt hàng"}
            </button>
          </div>
        </form>
      </Modal>

      {/* Transaction Modal */}
      <Modal 
        isOpen={isTransactionModalOpen} 
        onClose={() => setIsTransactionModalOpen(false)}
        title={editingId ? "CHỈNH SỬA GIAO DỊCH" : (activeTab === 'inbound' ? "NHẬP KHO MỚI" : "XUẤT KHO MỚI")}
      >
        <form onSubmit={handleAddTransaction} className="space-y-6">
          <div className="grid grid-cols-2 gap-4">
            <div className="col-span-2">
              <label className="block text-[10px] font-bold uppercase mb-2">Chọn Sản Phẩm</label>
              <select 
                required 
                value={newTransaction.productId}
                onChange={e => {
                  const p = products.find(p => p.id === e.target.value);
                  setNewTransaction({
                    ...newTransaction, 
                    productId: e.target.value,
                    lotNo: p?.lotNo || '',
                    ghiChu: p?.ghiChu || '',
                    designationCode: p?.designationCode || '',
                    loaiChiDinh: p?.loaiChiDinh || ''
                  });
                }}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none"
              >
                <option value="">-- Chọn mặt hàng --</option>
                {products.map(p => (
                  <option key={p.id} value={p.id}>{p.sku} - {p.name}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Số Lượng</label>
              <input 
                type="number" required 
                value={newTransaction.quantity}
                onChange={e => setNewTransaction({...newTransaction, quantity: Number(e.target.value)})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none font-bold" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Ngày Thực Hiện</label>
              <input 
                type="text" required 
                value={newTransaction.date}
                onChange={e => setNewTransaction({...newTransaction, date: e.target.value})}
                placeholder="dd/mm/yyyy"
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">{activeTab === 'inbound' ? 'Nhà cung cấp' : 'Khách hàng / Đối tác'}</label>
              <input 
                type="text" required 
                value={newTransaction.partner}
                onChange={e => setNewTransaction({...newTransaction, partner: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Lot No</label>
              <input 
                type="text" 
                value={newTransaction.lotNo}
                onChange={e => setNewTransaction({...newTransaction, lotNo: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Loại chỉ định</label>
              <input 
                type="text" 
                value={newTransaction.loaiChiDinh}
                onChange={e => setNewTransaction({...newTransaction, loaiChiDinh: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Mã chỉ định</label>
              <input 
                type="text" 
                value={newTransaction.designationCode}
                onChange={e => setNewTransaction({...newTransaction, designationCode: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
          </div>
          <div>
            <label className="block text-[10px] font-bold uppercase mb-2">Ghi Chú</label>
            <textarea 
              value={newTransaction.ghiChu}
              onChange={e => setNewTransaction({...newTransaction, ghiChu: e.target.value})}
              className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              rows={2}
            />
          </div>
          <div className="flex justify-end gap-3 pt-6">
            <button 
              type="button" 
              onClick={() => setIsTransactionModalOpen(false)}
              className="px-6 py-3 border border-[#141414] text-xs font-bold uppercase tracking-widest hover:bg-gray-100 transition-all"
            >
              Hủy
            </button>
            <button 
              type="submit"
              className="px-8 py-3 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-widest hover:opacity-90 transition-all"
            >
              {editingId ? "Cập nhật" : (activeTab === 'inbound' ? "Lưu Nhập Kho" : "Lưu Xuất Kho")}
            </button>
          </div>
        </form>
      </Modal>

      {/* Customer Modal */}
      <Modal 
        isOpen={isCustomerModalOpen} 
        onClose={() => setIsCustomerModalOpen(false)}
        title={editingId ? "CHỈNH SỬA KHÁCH HÀNG" : "THÊM KHÁCH HÀNG MỚI"}
      >
        <form onSubmit={handleAddCustomer} className="space-y-6">
          <div className="space-y-4">
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Mã Khách Hàng</label>
              <input 
                type="text" required 
                value={newCustomer.code}
                onChange={e => setNewCustomer({...newCustomer, code: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Tên Khách Hàng</label>
              <input 
                type="text" required 
                value={newCustomer.name}
                onChange={e => setNewCustomer({...newCustomer, name: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
          </div>
          <div className="flex justify-end gap-3 pt-6">
            <button 
              type="button" 
              onClick={() => setIsCustomerModalOpen(false)}
              className="px-6 py-3 border border-[#141414] text-xs font-bold uppercase tracking-widest hover:bg-gray-100 transition-all"
            >
              Hủy
            </button>
            <button 
              type="submit"
              className="px-8 py-3 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-widest hover:opacity-90 transition-all"
            >
              {editingId ? "Cập nhật" : "Lưu khách hàng"}
            </button>
          </div>
        </form>
      </Modal>

      {/* Location Modal */}
      <Modal 
        isOpen={isLocationModalOpen} 
        onClose={() => setIsLocationModalOpen(false)}
        title={editingId ? "CHỈNH SỬA VỊ TRÍ" : "THÊM VỊ TRÍ MỚI"}
      >
        <form onSubmit={handleLocationSubmit} className="space-y-6">
          <div className="grid grid-cols-2 gap-4">
            <div className="col-span-2">
              <label className="block text-[10px] font-bold uppercase mb-2">Mã QR (QRCode)</label>
              <input 
                type="text" required 
                value={newLocationEntry.qrcode}
                onChange={e => {
                  const qrcode = e.target.value;
                  const parsed = parseQRCode(qrcode);
                  setNewLocationEntry({
                    ...newLocationEntry,
                    qrcode,
                    sku: parsed?.sku || newLocationEntry.sku,
                    partner: parsed?.partner || newLocationEntry.partner,
                    date: parsed?.date || newLocationEntry.date
                  });
                }}
                className="w-full p-3 border border-[#141414] text-xs font-mono focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Mã SP (SKU)</label>
              <input 
                type="text" required 
                value={newLocationEntry.sku}
                onChange={e => setNewLocationEntry({...newLocationEntry, sku: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Nhà Cung Cấp</label>
              <input 
                type="text" 
                value={newLocationEntry.partner}
                onChange={e => setNewLocationEntry({...newLocationEntry, partner: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Ngày</label>
              <input 
                type="text" 
                value={newLocationEntry.date}
                onChange={e => setNewLocationEntry({...newLocationEntry, date: e.target.value})}
                placeholder="dd/mm/yyyy"
                className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Vị Trí</label>
              <input 
                type="text" required 
                value={newLocationEntry.location}
                onChange={e => setNewLocationEntry({...newLocationEntry, location: e.target.value})}
                className="w-full p-3 border border-[#141414] text-xs font-bold focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-2">Số lượng (Cuộn)</label>
              <input 
                type="number" 
                value={newLocationEntry.quantity}
                onChange={e => setNewLocationEntry({...newLocationEntry, quantity: parseInt(e.target.value) || 1})}
                className="w-full p-3 border border-[#141414] text-xs font-bold focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
          </div>
          <div>
            <label className="block text-[10px] font-bold uppercase mb-2">Ghi Chú</label>
            <textarea 
              value={newLocationEntry.note}
              onChange={e => setNewLocationEntry({...newLocationEntry, note: e.target.value})}
              className="w-full p-3 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              rows={2}
            />
          </div>
          <div className="flex justify-end gap-3 pt-6">
            <button 
              type="button" 
              onClick={() => setIsLocationModalOpen(false)}
              className="px-6 py-3 border border-[#141414] text-xs font-bold uppercase tracking-widest hover:bg-gray-100 transition-all"
            >
              Hủy
            </button>
            <button 
              type="submit"
              className="px-8 py-3 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-widest hover:opacity-90 transition-all"
            >
              {editingId ? "Cập nhật" : "Lưu vị trí"}
            </button>
          </div>
        </form>
      </Modal>

      {/* Delete Confirmation Modal */}
      <Modal 
        isOpen={isDeleteConfirmOpen} 
        onClose={() => setIsDeleteConfirmOpen(false)}
        title="XÁC NHẬN XÓA"
      >
        <div className="space-y-6">
          <div className="flex items-center gap-4 text-red-600">
            <AlertTriangle size={32} />
            <div>
              <p className="text-sm font-bold uppercase">Bạn có chắc chắn muốn xóa?</p>
              <p className="text-xs opacity-60">Hành động này không thể hoàn tác.</p>
            </div>
          </div>
          <div className="flex justify-end gap-3 pt-6">
            <button 
              onClick={() => setIsDeleteConfirmOpen(false)}
              className="px-6 py-3 border border-[#141414] text-xs font-bold uppercase tracking-widest hover:bg-gray-100 transition-all"
            >
              Hủy
            </button>
            <button 
              onClick={confirmDelete}
              className="px-8 py-3 bg-red-600 text-white text-xs font-bold uppercase tracking-widest hover:bg-red-700 transition-all"
            >
              Xác nhận xóa
            </button>
          </div>
        </div>
      </Modal>

      {/* Delivery Note Delete Confirmation Modal */}
      <Modal 
        isOpen={isDeliveryNoteDeleteConfirmOpen} 
        onClose={() => setIsDeliveryNoteDeleteConfirmOpen(false)}
        title="XÓA MỤC PHIẾU GIAO NHẬN"
      >
        <div className="space-y-6">
          <div className="flex items-center gap-4 text-red-600">
            <AlertTriangle size={32} />
            <div>
              <p className="text-sm font-bold uppercase">Xác nhận xóa dòng này khỏi phiếu?</p>
              <p className="text-xs opacity-60">Dữ liệu sẽ bị xóa khỏi danh sách đang soạn.</p>
            </div>
          </div>
          <div className="flex justify-end gap-3 pt-6">
            <button 
              onClick={() => setIsDeliveryNoteDeleteConfirmOpen(false)}
              className="px-6 py-3 border border-[#141414] text-xs font-bold uppercase tracking-widest hover:bg-gray-100 transition-all"
            >
              Hủy
            </button>
            <button 
              onClick={confirmDeleteDeliveryNoteItem}
              className="px-8 py-3 bg-red-600 text-white text-xs font-bold uppercase tracking-widest hover:bg-red-700 transition-all"
            >
              Xóa dòng
            </button>
          </div>
        </div>
      </Modal>

      {/* Delivery Note Edit Modal */}
      <Modal 
        isOpen={isDeliveryNoteEditModalOpen} 
        onClose={() => setIsDeliveryNoteEditModalOpen(false)}
        title="CHỈNH SỬA DÒNG LỆNH"
      >
        <form onSubmit={saveDeliveryNoteItemEdit} className="space-y-6">
          <div className="grid grid-cols-2 gap-4">
            <div className="col-span-2">
              <label className="block text-[10px] font-bold uppercase mb-1">Tên Hàng Hóa</label>
              <input 
                type="text" 
                value={tempDeliveryNoteItem.materialName}
                onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, materialName: e.target.value})}
                className="w-full p-2 border border-[#141414] text-xs focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-1">Mã Hàng (SKU)</label>
              <input 
                type="text" 
                value={tempDeliveryNoteItem.item}
                onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, item: e.target.value})}
                className="w-full p-2 border border-[#141414] text-xs font-mono focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-1">S.Lượng ERP</label>
              <input 
                type="number" 
                value={tempDeliveryNoteItem.qtyErp}
                onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, qtyErp: Number(e.target.value)})}
                className="w-full p-2 border border-[#141414] text-xs font-bold focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-1">OVN SO</label>
              <input 
                type="text" 
                value={tempDeliveryNoteItem.ovnSaleOrder}
                onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, ovnSaleOrder: e.target.value})}
                className="w-full p-2 border border-[#141414] text-xs font-mono focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-1">OVN Production Order</label>
              <input 
                type="text" 
                value={tempDeliveryNoteItem.ovnProductionOrder}
                onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, ovnProductionOrder: e.target.value})}
                className="w-full p-2 border border-[#141414] text-xs font-mono focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-1">No Code (Mã KH)</label>
              <input 
                type="text" 
                value={tempDeliveryNoteItem.noCode}
                onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, noCode: e.target.value})}
                className="w-full p-2 border border-[#141414] text-xs font-mono focus:ring-1 focus:ring-[#141414] outline-none" 
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold uppercase mb-1">Lot No</label>
              <input 
                type="text" 
                value={tempDeliveryNoteItem.lotNo}
                onChange={e => setTempDeliveryNoteItem({...tempDeliveryNoteItem, lotNo: e.target.value})}
                className="w-full p-2 border border-blue-600 text-xs font-bold bg-blue-50 focus:ring-1 focus:ring-blue-600 outline-none" 
                placeholder="Nhập hoặc để tự động..."
              />
            </div>
            <div className="col-span-2 p-3 bg-gray-50 border border-dashed border-[#141414]">
              <div className="flex justify-between items-center mb-1">
                <span className="text-[10px] font-bold uppercase opacity-60">Thông tin tồn kho gợi ý:</span>
                <span className="text-[10px] font-bold text-[#1a5f7a]">{tempDeliveryNoteItem.stock || 'Chưa tìm thấy tồn'}</span>
              </div>
              <div className="flex justify-between items-center">
                <span className="text-[10px] font-bold uppercase opacity-60">Vị trí gợi ý:</span>
                <span className="text-[10px] font-bold">{tempDeliveryNoteItem.location || 'N/A'}</span>
              </div>
            </div>
          </div>
          <div className="flex justify-end gap-3 pt-6">
            <button 
              type="button" 
              onClick={() => setIsDeliveryNoteEditModalOpen(false)}
              className="px-6 py-2 border border-[#141414] text-xs font-bold uppercase tracking-widest hover:bg-gray-100 transition-all"
            >
              Hủy
            </button>
            <button 
              type="submit"
              className="px-8 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-widest hover:opacity-90 transition-all"
            >
              Cập nhật dòng
            </button>
          </div>
        </form>
      </Modal>
    </div>
  );
}

// Sub-components
function NavItem({ active, onClick, icon, label }: { active: boolean, onClick: () => void, icon: React.ReactNode, label: string }) {
  return (
    <button 
      onClick={onClick}
      className={cn(
        "w-full flex items-center gap-4 px-4 py-3 text-xs font-bold uppercase tracking-widest transition-all",
        active 
          ? "bg-[#141414] text-[#E4E3E0] shadow-lg translate-x-1" 
          : "text-[#141414]/60 hover:bg-[#141414]/5 hover:text-[#141414]"
      )}
    >
      <span className={cn("transition-transform", active && "scale-110")}>{icon}</span>
      <span className="line-clamp-1">{label}</span>
    </button>
  );
}

function StatCard({ label, value, bgColor, textColor }: { label: string, value: number, bgColor: string, textColor: string }) {
  return (
    <div className={cn("p-6 border border-[#141414] flex flex-col justify-between h-32 sm:h-40 transition-transform hover:-translate-y-1", bgColor)}>
      <p className={cn("text-[9px] sm:text-[10px] font-bold uppercase tracking-[0.2em]", textColor)}>{label}</p>
      <p className={cn("text-3xl sm:text-4xl font-bold tracking-tighter", textColor)}>{value}</p>
    </div>
  );
}

function Modal({ isOpen, onClose, title, children }: { isOpen: boolean, onClose: () => void, title: string, children: React.ReactNode }) {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        className="bg-[#E4E3E0] w-full max-w-2xl border-2 border-[#141414] shadow-2xl overflow-hidden flex flex-col max-h-[90vh]"
      >
        <div className="p-6 border-b border-[#141414] flex justify-between items-center bg-[#141414] text-[#E4E3E0]">
          <h3 className="text-xs font-bold uppercase tracking-widest">{title}</h3>
          <button onClick={onClose} className="p-1 hover:rotate-90 transition-transform">
            <X size={20} />
          </button>
        </div>
        <div className="p-8 overflow-y-auto">
          {children}
        </div>
      </motion.div>
    </div>
  );
}
