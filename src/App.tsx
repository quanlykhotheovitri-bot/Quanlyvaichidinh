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
  X
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
    date: format(new Date(), 'dd/MM/yyyy')
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
    qrcode: '', sku: '', partner: '', date: format(new Date(), 'dd/MM/yyyy'), location: '', note: '', quantity: 1
  });
  const [deleteTarget, setDeleteTarget] = useState<{ id: string | 'bulk', type: 'product' | 'transaction' | 'customer' | 'location' | 'savedDeliveryNote', qrcode?: string } | null>(null);
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

  const generateId = useCallback(() => {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) {
      return crypto.randomUUID();
    }
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
      var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }, []);

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

  const saveDeliveryNoteItemEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (editingDeliveryNoteId !== null) {
      const updatedNotes = deliveryNotes.map(item => 
        item.id === editingDeliveryNoteId ? { ...item, ...tempDeliveryNoteItem } as DeliveryNoteItem : item
      );
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
    setDeliveryNotes(prev => prev.map(item => 
      item.id === itemAtId ? { ...item, [field]: value } : item
    ));
  };

  const handlePostDeliveryNote = async () => {
    if (deliveryNotes.length === 0) return;

    const today = format(new Date(), 'dd/MM/yyyy');
    const newTransactions: Transaction[] = [];

    deliveryNotes.forEach(item => {
      if (item.actualQty > 0) {
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
      updatedEntry = { ...newLocationEntry, id: editingId, type: locationSubTab === 'inventory' ? 'inventory' : 'input' } as LocationEntry;
      if (locationSubTab === 'input' || locationSubTab === 'output') {
        setLocationEntries(locationEntries.map(entry => entry.id === editingId ? updatedEntry : entry));
      } else {
        setLocationInventoryEntries(locationInventoryEntries.map(entry => entry.id === editingId ? updatedEntry : entry));
      }
      setEditingId(null);
    } else {
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
            date: newLocationEntry.date || existing.date,
            note: newLocationEntry.note || existing.note,
            quantity: (existing.quantity || 0) + (newLocationEntry.quantity || 0),
          };
          setLocationInventoryEntries(prev => prev.map((e, i) => i === existingIndex ? updatedEntry : e));
          setEditingId(null);
        } else {
          updatedEntry = {
            ...newLocationEntry as LocationEntry,
            id: generateId(),
            type: locationSubTab === 'inventory' ? 'inventory' : 'input'
          };
          setLocationInventoryEntries(prev => [...prev, updatedEntry]);
        }
      } else {
        updatedEntry = {
          ...newLocationEntry as LocationEntry,
          id: generateId(),
          type: 'input',
          scanType: scanMode
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
              type: 'inventory'
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
    setNewLocationEntry({ qrcode: '', sku: '', partner: '', date: format(new Date(), 'dd/MM/yyyy'), location: '', note: '' });
  };

  const processLocationData = async (data: any[]) => {
    const newEntries: any[] = data.map((row: any) => {
      const normalizedRow: any = {};
      Object.keys(row).forEach(key => {
        normalizedRow[key.toLowerCase().trim()] = row[key];
      });

      const qrcode = String(normalizedRow['qrcode'] || normalizedRow['qr code'] || row['QR Code'] || row['qrcode'] || '').trim();
      const parsed = parseQRCode(qrcode);
      
      const isSimpleAWB = qrcode.toUpperCase().startsWith('AWB-') && !qrcode.includes('|');
      const defaultDate = isSimpleAWB ? '' : format(new Date(), 'dd/MM/yyyy');

      return {
        qrcode,
        sku: normalizedRow['sku'] || row['SKU'] || row['sku'] || parsed?.sku || '',
        partner: normalizedRow['partner'] || normalizedRow['đối tác'] || row['Đối tác'] || row['partner'] || parsed?.partner || '',
        date: normalizedRow['date'] || normalizedRow['ngày'] || row['Ngày'] || row['date'] || parsed?.date || defaultDate,
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

  const confirmDelete = async () => {
    if (!deleteTarget) return;
    if (deleteTarget.id === 'bulk') {
      await handleBulkDelete();
    } else {
      await handleDelete(deleteTarget.id, deleteTarget.type);
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

      const batchKey = `${t.productId}-${t.lotNo || ''}-${t.designationCode || ''}-${t.loaiChiDinh || ''}`;
      
      if (!batches[batchKey]) {
        batches[batchKey] = {
          ...product,
          id: batchKey,
          productId: product.id,
          lotNo: t.lotNo || '',
          ghiChu: t.ghiChu || '',
          designationCode: t.designationCode || '',
          loaiChiDinh: t.loaiChiDinh || '',
          totalInbound: 0,
          totalOutbound: 0,
          currentStock: 0
        };
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
          currentStock: 0
        };
      }
    });

    return Object.values(batches);
  }, [products, transactions]);

  const filteredInventory = useMemo(() => {
    const query = searchQuery.toLowerCase();
    return inventory.filter(item => 
      item.name.toLowerCase().includes(query) ||
      item.sku.toLowerCase().includes(query)
    );
  }, [inventory, searchQuery]);

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
    if (activeTab === 'inventory') {
      setProducts(prev => {
        const updatedProducts = [...prev];
        const skuToProductIndex = new Map(updatedProducts.map((p, i) => [p.sku, i]));

        data.forEach((row, index) => {
          const normalizedRow: any = {};
          Object.keys(row).forEach(key => {
            normalizedRow[key.toLowerCase().trim()] = row[key];
          });

          const sku = String(
            normalizedRow['sku'] || 
            normalizedRow['mã hàng'] || 
            normalizedRow['mã sp'] || 
            normalizedRow['mã sản phẩm'] || 
            normalizedRow['mã vật tư'] ||
            normalizedRow['item no'] ||
            row.sku || row.SKU || row['Mã Hàng'] || ''
          ).trim();

          if (!sku) return;

          const name = String(
            normalizedRow['name'] || 
            normalizedRow['tên hàng'] || 
            normalizedRow['tên sp'] || 
            normalizedRow['tên sản phẩm'] || 
            normalizedRow['tên vật tư'] ||
            normalizedRow['product name'] || 
            normalizedRow['item name'] ||
            row.name || row.Name || row['Tên hàng'] || 'Sản phẩm mới'
          ).trim();

          const existingIndex = skuToProductIndex.get(sku);
          const productData: Product = {
            id: existingIndex !== undefined ? updatedProducts[existingIndex].id : generateId(),
            sku: sku,
            name: name,
            category: normalizedRow['category'] || normalizedRow['loại'] || row.category || row.Category || 'Chưa phân loại',
            unit: normalizedRow['unit'] || normalizedRow['đơn vị'] || row.unit || row.Unit || 'Cái',
            minStock: Number(normalizedRow['minstock'] || normalizedRow['tồn tối thiểu'] || row.minStock || row.MinStock || 0),
            lotNo: normalizedRow['lotno'] || normalizedRow['lot no'] || row.lotNo || row.LotNo || row['Lot no'] || '',
            ghiChu: normalizedRow['ghichu'] || normalizedRow['ghi chú'] || row.ghiChu || row.GhiChu || row['Ghi chú'] || '',
            designationCode: normalizedRow['designationcode'] || normalizedRow['mã chỉ định'] || row.designationCode || row.DesignationCode || row['Mã chỉ định'] || '',
            loaiChiDinh: normalizedRow['loaichidinh'] || normalizedRow['loại chỉ định'] || row.loaiChiDinh || row.LoaiChiDinh || row['Loại chỉ định'] || ''
          };

          if (existingIndex !== undefined) {
            updatedProducts[existingIndex] = productData;
          } else {
            skuToProductIndex.set(sku, updatedProducts.length);
            updatedProducts.push(productData);
          }
        });
        api.products.upsertAll(updatedProducts).catch(err => console.error('Error syncing products:', err));
        return updatedProducts;
      });
    } else if (activeTab === 'inbound' || activeTab === 'outbound') {
      setProducts(prevProducts => {
        const updatedProducts = [...prevProducts];
        const newProductsToAdd: Product[] = [];
        
        const skuToProductIndex = new Map(updatedProducts.map((p, i) => [p.sku, i]));
        const skuToNewProductIndex = new Map<string, number>();

        data.forEach((row, index) => {
          const normalizedRow: any = {};
          Object.keys(row).forEach(key => {
            normalizedRow[key.toLowerCase().trim()] = row[key];
          });

          const sku = String(
            normalizedRow['sku'] || 
            normalizedRow['mã hàng'] || 
            normalizedRow['mã sp'] || 
            normalizedRow['mã sản phẩm'] || 
            normalizedRow['mã vật tư'] ||
            normalizedRow['item no'] ||
            row.sku || row.SKU || row['Mã Hàng'] || ''
          ).trim();

          if (!sku) return;

          const name = String(
            normalizedRow['name'] || 
            normalizedRow['tên hàng'] || 
            normalizedRow['tên sp'] || 
            normalizedRow['tên sản phẩm'] || 
            normalizedRow['tên vật tư'] ||
            normalizedRow['product name'] || 
            normalizedRow['item name'] ||
            row.name || row.Name || row['Tên hàng'] || ''
          ).trim();

          const designationCode = String(
            normalizedRow['designationcode'] || 
            normalizedRow['mã chỉ định'] || 
            row.designationCode || row.DesignationCode || row['Mã chỉ định'] || ''
          ).trim();

          const loaiChiDinh = String(
            normalizedRow['loaichidinh'] || 
            normalizedRow['loại chỉ định'] || 
            row.loaiChiDinh || row.LoaiChiDinh || row['Loại chỉ định'] || ''
          ).trim();

          const ghiChu = String(
            normalizedRow['ghichu'] || 
            normalizedRow['ghi chú'] || 
            row.ghiChu || row.GhiChu || row['Ghi chú'] || ''
          ).trim();

          const existingIndex = skuToProductIndex.get(sku);
          const alreadyInNewIndex = skuToNewProductIndex.get(sku);
          
          if (existingIndex !== undefined) {
            if (name && updatedProducts[existingIndex].name !== name && name !== 'Sản phẩm mới') {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], name };
            }
            if (designationCode && updatedProducts[existingIndex].designationCode !== designationCode) {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], designationCode };
            }
            if (loaiChiDinh && updatedProducts[existingIndex].loaiChiDinh !== loaiChiDinh) {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], loaiChiDinh };
            }
            if (ghiChu && updatedProducts[existingIndex].ghiChu !== ghiChu) {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], ghiChu };
            }
          } else if (alreadyInNewIndex !== undefined) {
            if (name && newProductsToAdd[alreadyInNewIndex].name === 'Sản phẩm mới' && name !== 'Sản phẩm mới') {
              newProductsToAdd[alreadyInNewIndex] = { ...newProductsToAdd[alreadyInNewIndex], name };
            }
            if (designationCode && newProductsToAdd[alreadyInNewIndex].designationCode === '') {
              newProductsToAdd[alreadyInNewIndex] = { ...newProductsToAdd[alreadyInNewIndex], designationCode };
            }
            if (loaiChiDinh && (!newProductsToAdd[alreadyInNewIndex].loaiChiDinh || newProductsToAdd[alreadyInNewIndex].loaiChiDinh === '')) {
              newProductsToAdd[alreadyInNewIndex] = { ...newProductsToAdd[alreadyInNewIndex], loaiChiDinh };
            }
            if (ghiChu && newProductsToAdd[alreadyInNewIndex].ghiChu === '') {
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
              unit: normalizedRow['unit'] || normalizedRow['đơn vị'] || 'Cái',
              minStock: 0,
              lotNo: normalizedRow['lotno'] || normalizedRow['lot no'] || '',
              ghiChu: ghiChu || '',
              designationCode: designationCode || '',
              loaiChiDinh: loaiChiDinh || ''
            });
          }
        });

        const finalProducts = [...updatedProducts, ...newProductsToAdd];
        const finalProductsMap = new Map(finalProducts.map(p => [p.sku, p]));

        const newTransactions: Transaction[] = data.map((row, index) => {
          const normalizedRow: any = {};
          Object.keys(row).forEach(key => {
            normalizedRow[key.toLowerCase().trim()] = row[key];
          });

          const sku = String(
            normalizedRow['sku'] || 
            normalizedRow['mã hàng'] || 
            normalizedRow['mã sp'] || 
            normalizedRow['mã sản phẩm'] || 
            normalizedRow['mã vật tư'] ||
            normalizedRow['item no'] ||
            row.sku || row.SKU || row['Mã Hàng'] || ''
          ).trim();

          if (!sku) return null;

          const product = finalProductsMap.get(sku);
          if (!product) return null;

          const rawDate = normalizedRow['date'] || normalizedRow['ngày'] || normalizedRow['ngày nhập'] || normalizedRow['ngày xuất'] || row.date || row.Date || row['Ngày nhập'] || row['Ngày xuất'];
          let formattedDate = format(new Date(), 'dd/MM/yyyy');
          
          if (rawDate) {
            if (typeof rawDate === 'number') {
              const date = XLSX.SSF.parse_date_code(rawDate);
              formattedDate = format(new Date(date.y, date.m - 1, date.d), 'dd/MM/yyyy');
            } else {
              formattedDate = String(rawDate).trim();
            }
          }

          return {
            id: generateId(),
            productId: product.id,
            type: activeTab as 'inbound' | 'outbound',
            quantity: Number(normalizedRow['quantity'] || normalizedRow['số lượng'] || normalizedRow['số lượng nhập'] || normalizedRow['số lượng xuất'] || row.quantity || row.Quantity || row['Số lượng nhập'] || row['Số lượng xuất'] || 0),
            date: formattedDate,
            partner: normalizedRow['partner'] || normalizedRow['đối tác'] || normalizedRow['khách hàng'] || normalizedRow['nhà cung cấp'] || row.partner || row.Partner || 'N/A',
            loaiChiDinh: normalizedRow['loaichidinh'] || normalizedRow['loại chỉ định'] || row.loaiChiDinh || row.LoaiChiDinh || row['Loại chỉ định'] || '',
            lotNo: normalizedRow['lotno'] || normalizedRow['lot no'] || row.lotNo || row.LotNo || row['Lot no'] || product?.lotNo || '',
            ghiChu: normalizedRow['ghichu'] || normalizedRow['ghi chú'] || row.ghiChu || row.GhiChu || row['Ghi chú'] || product?.ghiChu || '',
            designationCode: normalizedRow['designationcode'] || normalizedRow['mã chỉ định'] || row.designationCode || row.DesignationCode || row['Mã chỉ định'] || product?.designationCode || ''
          };
        }).filter(Boolean) as Transaction[];

        if (newTransactions.length > 0) {
          setTransactions(prev => [...prev, ...newTransactions]);
          api.transactions.upsertAll(newTransactions).catch(err => console.error('Error syncing transactions:', err));
        }

        api.products.upsertAll(finalProducts).catch(err => console.error('Error syncing products:', err));
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
        Promise.all(newCustomers.map(c => api.customers.upsert(c))).catch(err => console.error('Error syncing customers:', err));
        return updated;
      });
    }
  };

  const parseQRCode = (qrcode: string) => {
    if (!qrcode) return null;
    const parts = qrcode.split('|');
    
    if (parts.length < 2) {
      return { sku: qrcode.trim(), partner: '', date: '' };
    }
    
    const sku = parts[0].trim();
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
                  await Promise.all(entriesToDelete.map(e => api.locationEntries.delete(e.id)));
                }
                if (finalEntries.length > 0) {
                  await api.locationEntries.upsertAll(finalEntries);
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
          const jsonData = XLSX.utils.sheet_to_json(worksheet);
          
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
    const productMap = new Map<string, Product>(products.map(p => [p.sku, p]));
    const customerMap = new Map<string, Customer>(customers.map(c => [c.name.toLowerCase(), c]));
    
    const inventoryBySku = new Map<string, InventoryItem[]>();
    inventory.forEach(item => {
      if (!inventoryBySku.has(item.sku)) {
        inventoryBySku.set(item.sku, []);
      }
      inventoryBySku.get(item.sku)!.push(item);
    });

    const priorityOrder = [
      "CHI DINH THEO SO",
      "CHI DINH theo KH",
      "CHI DINH THEO NCC",
      "NORMAL",
      "VAI KH CUNG CAP"
    ];

    const extractDateFromLot = (lot: string): number => {
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
    };

    const normalizeDateForMatching = (lot: string): string => {
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
    };

    const workingInventory = inventory.map(item => ({
      ...item,
      tempStock: item.currentStock
    }));

    const mappedData: DeliveryNoteItem[] = data.map((row, index) => {
      const itemNo = String(row['Item No.'] || row['item'] || '').trim();
      const product = productMap.get(itemNo);
      const customerName = String(row['Sell-to Customer Name'] || row['Customer code'] || '').trim();
      const customer = customerMap.get(customerName.toLowerCase());
      
      const ovnSaleOrder = String(row['OVN Sale Order'] || '').trim();
      const ovnProductionOrder = String(row['OVN Production Order'] || '').trim();
      const qtyNeeded = Number(row['Quantity'] || row['Qty ERP'] || 0);

      const availableStock = workingInventory.filter(item => item.sku === itemNo && item.tempStock > 0);
      
      let bestBatch: typeof workingInventory[0] | undefined;
      
      for (const priorityType of priorityOrder) {
        const matchingBatches = availableStock.filter(item => 
          (item.loaiChiDinh || '').trim().toUpperCase() === priorityType.toUpperCase()
        );
        
        if (matchingBatches.length > 0) {
          matchingBatches.sort((a, b) => extractDateFromLot(a.lotNo || '') - extractDateFromLot(b.lotNo || ''));

          if (priorityType === "CHI DINH THEO SO") {
            const orderMatch = matchingBatches.find(item => 
              (item.designationCode || '').trim() === ovnSaleOrder || 
              (item.designationCode || '').trim() === ovnProductionOrder
            );
            if (orderMatch) {
              bestBatch = orderMatch;
              break;
            }
          } else if (priorityType === "CHI DINH theo KH") {
            const customerMatch = matchingBatches.find(item => {
              const designationCode = (item.designationCode || '').trim();
              const customerCode = (customer?.code || '').trim();
              if (!customerCode) return false;
              
              const codes = designationCode.split('/').map(c => c.trim());
              return codes.includes(customerCode);
            });
            if (customerMatch) {
              bestBatch = customerMatch;
              break;
            }
          } else {
            bestBatch = matchingBatches[0];
            break;
          }
        }
      }

      if (bestBatch) {
        bestBatch.tempStock -= qtyNeeded;
      }

      const normalizedLotDate = normalizeDateForMatching(bestBatch?.lotNo || '');
      const locationEntry = locationInventoryEntries.find(e => 
        e.sku === itemNo && e.date === normalizedLotDate
      );

      return {
        id: generateId(),
        no: 0,
        ovnSaleOrder: ovnSaleOrder,
        ovnProductionOrder: ovnProductionOrder,
        item: itemNo,
        materialName: String(row['OVN Full Name'] || product?.name || ''),
        unit: 'YDS',
        qtyErp: qtyNeeded,
        actualQty: bestBatch ? qtyNeeded : 0,
        lotNo: bestBatch?.lotNo || '',
        remark: bestBatch?.designationCode || '',
        brand: String(row['Brand Code'] || ''),
        customerCode: customerName,
        finalDestination: String(row['Final Destination'] || ''),
        noCode: customer?.code || '',
        location: locationEntry ? locationEntry.location : (bestBatch ? 'Chưa có vị trí' : ''),
        stock: bestBatch ? `${bestBatch.currentStock} (${bestBatch.loaiChiDinh})` : 'Không có tồn'
      };
    });

    const sortedData = mappedData.sort((a, b) => a.item.localeCompare(b.item));
    
    const groups: { [key: string]: DeliveryNoteItem[] } = {};
    sortedData.forEach(item => {
      if (!groups[item.item]) groups[item.item] = [];
      groups[item.item].push(item);
    });

    Object.values(groups).forEach(group => {
      const validRows = group.filter(item => item.stock !== 'Không có tồn');

      if (validRows.length > 0) {
        const totalQtyErp = validRows.reduce((sum, item) => sum + item.qtyErp, 0);
        const roundedTotal = Math.ceil(totalQtyErp);
        const diff = roundedTotal - totalQtyErp;
        
        let maxRow = validRows[0];
        validRows.forEach(item => {
          if (item.qtyErp > maxRow.qtyErp) maxRow = item;
        });
        
        maxRow.actualQty = maxRow.qtyErp + diff;
        
        group.forEach(item => {
          item.actualIssuedQty = roundedTotal;
        });
      } else {
        group.forEach(item => {
          item.actualQty = 0;
          item.actualIssuedQty = 0;
        });
      }
    });

    const finalData = sortedData.map((item, index) => ({
      ...item,
      no: index + 1
    }));

    setDeliveryNotes(finalData);
    api.deliveryNotes.upsertAll(finalData).catch(err => console.error('Error syncing delivery notes:', err));
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
    <div className="min-h-screen bg-[#E4E3E0] text-[#141414] font-sans flex">
      <aside className="w-64 border-r border-[#141414] flex flex-col">
        <div className="p-6 border-bottom border-[#141414]">
          <h1 className="font-serif italic text-2xl font-bold tracking-tight">KHO.LOG</h1>
          <p className="text-[10px] uppercase tracking-widest opacity-50 mt-1">Warehouse Management System</p>
        </div>

        <nav className="flex-1 px-4 py-8 space-y-2">
          <NavItem 
            active={activeTab === 'dashboard'} 
            onClick={() => setActiveTab('dashboard')}
            icon={<LayoutDashboard size={18} />}
            label="Tổng quan"
          />
          <NavItem 
            active={activeTab === 'inbound'} 
            onClick={() => setActiveTab('inbound')}
            icon={<ArrowDownToLine size={18} />}
            label="Nhập kho"
          />
          <NavItem 
            active={activeTab === 'outbound'} 
            onClick={() => setActiveTab('outbound')}
            icon={<ArrowUpFromLine size={18} />}
            label="Xuất kho"
          />
          <NavItem 
            active={activeTab === 'inventory'} 
            onClick={() => setActiveTab('inventory')}
            icon={<Package size={18} />}
            label="Tồn kho"
          />
          <NavItem 
            active={activeTab === 'customers'} 
            onClick={() => setActiveTab('customers')}
            icon={<Users size={18} />}
            label="Danh sách khách hàng"
          />
          <NavItem 
            active={activeTab === 'deliveryNote'} 
            onClick={() => setActiveTab('deliveryNote')}
            icon={<FileText size={18} />}
            label="Lệnh xuất kho"
          />
          <NavItem 
            active={activeTab === 'location'} 
            onClick={() => setActiveTab('location')}
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

      <main className="flex-1 flex flex-col overflow-hidden">
        <header className="h-20 border-b border-[#141414] flex items-center justify-between px-8 bg-[#E4E3E0]/80 backdrop-blur-sm z-10">
          <div className="flex items-center gap-4 flex-1 max-w-md">
            <Search size={18} className="opacity-40" />
            <input 
              type="text" 
              placeholder="Tìm kiếm sản phẩm, mã SKU, đối tác..." 
              className="bg-transparent border-none outline-none w-full text-sm placeholder:italic"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
            />
          </div>

          <div className="flex items-center gap-4">
            <button 
              onClick={handleSaveAll}
              disabled={isSaving}
              className={`flex items-center gap-2 px-4 py-2 text-[10px] font-bold uppercase tracking-wider transition-all disabled:opacity-80 ${!isOnline ? 'bg-red-600 text-white' : isSaving ? 'bg-[#ff9900] text-black' : 'bg-[#141414] text-[#E4E3E0] hover:opacity-90'}`}
              title={!isOnline ? "Hoạt động Ngoại Tuyến (Offline)" : "Hệ thống đã bật tự động lưu"}
            >
              {!isOnline ? (
                <WifiOff size={14} />
              ) : isSaving ? (
                <div className="w-3 h-3 border-2 border-black border-t-transparent rounded-full animate-spin" />
              ) : (
                <Save size={14} />
              )}
              {isSaving ? 'Đang lưu...' : 'Lưu tất cả'}
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
                className="flex items-center gap-2 px-4 py-2 bg-red-600 text-white text-[10px] font-bold uppercase tracking-wider hover:bg-red-700 transition-colors"
              >
                <Trash2 size={14} />
                Xóa ({selectedRows.length})
              </button>
            )}
            {activeTab === 'deliveryNote' && deliveryNoteSubTab === 'preview' && (
              <button 
                onClick={() => setDeliveryNotes([])}
                className="flex items-center gap-2 px-3 py-2 border border-[#141414] text-[10px] font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
              >
                <Trash2 size={14} />
                Làm mới
              </button>
            )}
            <button 
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-2 px-4 py-2 border border-[#141414] text-[10px] font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
            >
              <FileUp size={14} />
              Cập nhật file
            </button>
            <input 
              type="file" 
              ref={fileInputRef} 
              onChange={handleFileUpload} 
              accept=".csv,.xlsx,.xls,.xlsm" 
              className="hidden" 
            />
            <button className="p-2 border border-[#141414] hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors">
              <Filter size={18} />
            </button>
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

        <div className="flex-1 overflow-y-auto p-8">
          <AnimatePresence mode="wait">
            {activeTab === 'dashboard' && (
              <motion.div 
                key="dashboard"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-8"
              >
                <div className="grid grid-cols-4 gap-6">
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

                <div className="grid grid-cols-4 gap-8">
                  <div className="col-span-3 border border-[#141414] p-6 bg-[#1a5f7a] text-white">
                    <div className="border border-white/30 p-2 inline-block mb-6">
                      <h3 className="text-xs font-bold uppercase tracking-widest">BIỂU ĐỒ TỒN KHO THEO LOẠI CHỈ ĐỊNH</h3>
                    </div>
                    <div className="h-[400px]">
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

                  <div className="border border-[#141414] p-8 bg-[#1a5f7a] text-white flex flex-col gap-4">
                    <div className="space-y-1">
                      <p className="text-[11px] font-bold uppercase tracking-wider opacity-80">TỔNG MÃ TỒN KHO :</p>
                      <p className="text-3xl font-bold">{stats.totalStockCodes.toLocaleString()}</p>
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
                <div className="flex items-center justify-between">
                  <h2 className="font-serif italic text-2xl capitalize">
                    {activeTab === 'inbound' ? 'Danh sách nhập kho' : 'Danh sách xuất kho'}
                  </h2>
                  <div className="flex gap-3">
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                      <input
                        type="text"
                        placeholder="Tìm kiếm giao dịch..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="pl-10 pr-4 py-2 border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-64"
                      />
                    </div>
                    <button 
                      onClick={() => {
                        setNewTransaction(prev => ({ ...prev, type: activeTab as 'inbound' | 'outbound' }));
                        setIsTransactionModalOpen(true);
                      }}
                      className="flex items-center gap-2 px-6 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <Plus size={14} />
                      Thêm giao dịch
                    </button>
                  </div>
                </div>

                <div className="border border-[#141414] overflow-x-auto">
                  <table className="w-full border-collapse min-w-[1000px]">
                    <thead>
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
                              <td className="border border-[#141414] p-3 italic">{t.lotNo || product?.lotNo}</td>
                              <td className="border border-[#141414] p-3 text-center font-bold">{t.quantity}</td>
                              <td className="border border-[#141414] p-3">{t.date}</td>
                              <td className="border border-[#141414] p-3 opacity-60">{t.loaiChiDinh}</td>
                              <td className="border border-[#141414] p-3">{t.ghiChu || product?.ghiChu}</td>
                              <td className="border border-[#141414] p-3 font-mono">{t.designationCode || product?.designationCode}</td>
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
                <div className="flex items-center justify-between">
                  <h2 className="font-serif italic text-2xl">Báo cáo tồn kho</h2>
                  <div className="flex gap-3">
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                      <input
                        type="text"
                        placeholder="Tìm kiếm tồn kho..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="pl-10 pr-4 py-2 border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-64"
                      />
                    </div>
                    <button className="flex items-center gap-2 px-6 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors">
                      <Download size={14} />
                      Xuất báo cáo
                    </button>
                    <button 
                      onClick={() => setIsProductModalOpen(true)}
                      className="flex items-center gap-2 px-6 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <Plus size={14} />
                      Thêm sản phẩm
                    </button>
                  </div>
                </div>

                <div className="border border-[#141414] overflow-x-auto">
                  <table className="w-full border-collapse min-w-[1200px]">
                    <thead>
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
                                    setNewProduct(item);
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
                <div className="flex items-center justify-between">
                  <h2 className="font-serif italic text-2xl">Danh sách khách hàng</h2>
                  <div className="flex gap-3">
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                      <input
                        type="text"
                        placeholder="Tìm kiếm khách hàng..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="pl-10 pr-4 py-2 border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-64"
                      />
                    </div>
                    <button 
                      onClick={() => fileInputRef.current?.click()}
                      className="flex items-center gap-2 px-6 py-2 border border-[#141414] text-xs font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                    >
                      <FileUp size={14} />
                      Tải lên file
                    </button>
                    <button 
                      onClick={() => {
                        setEditingId(null);
                        setNewCustomer({ code: '', name: '' });
                        setIsCustomerModalOpen(true);
                      }}
                      className="flex items-center gap-2 px-6 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <Plus size={14} />
                      Thêm khách hàng
                    </button>
                  </div>
                </div>

                <div className="border border-[#141414] overflow-x-auto">
                  <table className="w-full border-collapse">
                    <thead>
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
                <div className="flex justify-between items-center no-print">
                  <div className="flex items-center gap-6">
                    <h2 className="font-serif italic text-2xl">Lệnh xuất kho</h2>
                    <div className="flex border-b border-gray-200">
                      <button
                        onClick={() => setDeliveryNoteSubTab('preview')}
                        className={cn(
                          "px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors border-b-2",
                          deliveryNoteSubTab === 'preview' ? "border-[#141414] text-[#141414]" : "border-transparent text-gray-400 hover:text-gray-600"
                        )}
                      >
                        Phiếu giao nhận
                      </button>
                      <button
                        onClick={() => setDeliveryNoteSubTab('history')}
                        className={cn(
                          "px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors border-b-2",
                          deliveryNoteSubTab === 'history' ? "border-[#141414] text-[#141414]" : "border-transparent text-gray-400 hover:text-gray-600"
                        )}
                      >
                        Lịch sử đã Post
                      </button>
                    </div>
                  </div>
                  <div className="flex gap-3">
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={16} />
                      <input
                        type="text"
                        placeholder="Tìm kiếm lệnh xuất..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="pl-10 pr-4 py-2 border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-64"
                      />
                    </div>
                    {deliveryNoteSubTab === 'preview' && (
                      <>
                        <button 
                          onClick={handlePostDeliveryNote}
                          className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white text-[10px] font-bold uppercase tracking-wider hover:bg-green-700 transition-colors"
                        >
                          <CheckSquare size={14} />
                          Post xuất kho
                        </button>
                        <button 
                          onClick={() => {
                            window.print();
                          }}
                          className="flex items-center gap-2 px-4 py-2 border border-[#141414] text-[10px] font-bold uppercase tracking-wider hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
                        >
                          <Printer size={14} />
                          In phiếu
                        </button>
                      </>
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

                        const titleRow = worksheet.addRow(['', 'PHIẾU GIAO NHẬN FABRIC']);
                        titleRow.getCell(2).font = { name: 'Times New Roman', size: 16, bold: true };
                        titleRow.getCell(2).alignment = { horizontal: 'center' };
                        worksheet.mergeCells(1, 2, 1, 17);

                        const subtitleRow = worksheet.addRow(['', 'Delivery Note']);
                        subtitleRow.getCell(2).font = { name: 'Times New Roman', size: 12, italic: true };
                        subtitleRow.getCell(2).alignment = { horizontal: 'center' };
                        worksheet.mergeCells(2, 2, 2, 17);

                        worksheet.addRow([]);

                        const meta1 = worksheet.addRow(['Mã Tài Liệu:', deliveryNoteHeader.documentCode]);
                        meta1.getCell(1).font = { bold: true };
                        const meta2 = worksheet.addRow(['Dept:', deliveryNoteHeader.dept]);
                        meta2.getCell(1).font = { bold: true };
                        const meta3 = worksheet.addRow(['TO:', deliveryNoteHeader.to]);
                        meta3.getCell(1).font = { bold: true };
                        const meta4 = worksheet.addRow(['Date:', deliveryNoteHeader.date]);
                        meta4.getCell(1).font = { bold: true };

                        worksheet.addRow([]);

                        const headerRow = worksheet.addRow([
                          'No', 'OVN Sale Order', 'OVN Production Order', 'item', 'Material Name', 
                          'Unit', 'Qty ERP', 'Thực tế', 'Lot No', 'Số lượng thực phát', 
                          'remark', 'Brand', 'Customer code', 'Final Destination', 'No.', 'Vị trí', 'STOCK'
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
                          if (!processedGroups.has(item.item)) {
                            totalActualIssuedQty += (item.actualIssuedQty || 0);
                            processedGroups.add(item.item);
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
                            
                            if ([1, 6, 10].includes(colNumber)) {
                              cell.alignment = { horizontal: 'center' };
                            } else if ([7, 8].includes(colNumber)) {
                              cell.alignment = { horizontal: 'right' };
                            }
                          });
                        });

                        let currentItem = '';
                        let startMergeRow = 0;
                        const dataStartRow = 9;

                        deliveryNotes.forEach((item, index) => {
                          const rowIdx = dataStartRow + index;
                          if (item.item !== currentItem) {
                            if (startMergeRow !== 0 && (rowIdx - 1) > startMergeRow) {
                              worksheet.mergeCells(startMergeRow, 10, rowIdx - 1, 10);
                            }
                            currentItem = item.item;
                            startMergeRow = rowIdx;
                          }
                          if (index === deliveryNotes.length - 1) {
                            if (rowIdx > startMergeRow) {
                              worksheet.mergeCells(startMergeRow, 10, rowIdx, 10);
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
                        totalRow.eachCell((cell) => {
                          cell.font = { bold: true, size: 10 };
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
                  <div className="bg-white border border-[#141414] p-8 shadow-sm">
                  <div className="flex justify-center items-start mb-8">
                    <div className="text-center flex-1">
                      <h1 className="text-xl font-bold uppercase tracking-widest">Phiếu giao nhận Fabric</h1>
                      <p className="text-sm italic">Delivery Note</p>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-x-12 gap-y-2 mb-8 text-xs">
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

                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse min-w-[1800px]">
                      <thead>
                        <tr className="bg-[#001F3F] text-white text-[10px] uppercase tracking-wider">
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
                            <td colSpan={17} className="p-12 text-center text-gray-400 italic">
                              Chưa có dữ liệu. Vui lòng tải lên file nguồn để cập nhật.
                            </td>
                          </tr>
                        ) : (
                          filteredDeliveryNotes.map((item, index) => {
                            const isFirstInGroup = index === 0 || filteredDeliveryNotes[index - 1].item !== item.item;
                            const groupSize = isFirstInGroup ? filteredDeliveryNotes.filter(dn => dn.item === item.item).length : 0;

                            const isDesignationMatch = (remark: string, saleOrder: string, prodOrder: string, custCode: string) => {
                              if (!remark) return true;
                              const trimmedRemark = remark.trim();
                              if (trimmedRemark === saleOrder || trimmedRemark === prodOrder) return true;
                              if (!custCode) return false;
                              const codes = trimmedRemark.split('/').map(c => c.trim());
                              return codes.includes(custCode.trim());
                            };

                            const hasMismatch = !isDesignationMatch(item.remark, item.ovnSaleOrder, item.ovnProductionOrder, item.noCode);

                            return (
                              <tr 
                                key={item.id} 
                                className={cn(
                                  "text-[11px] transition-colors",
                                  !item.noCode ? "bg-yellow-200 text-red-600 font-bold" : "bg-white hover:bg-gray-50",
                                  hasMismatch ? "bg-red-100 text-red-700" : ""
                                )}
                              >
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
                                <td className="border border-[#141414] p-2 text-right font-mono">
                                  <input 
                                    type="number"
                                    value={item.actualQty || 0}
                                    onChange={(e) => handleEditDeliveryNoteItem(index, 'actualQty', Number(e.target.value))}
                                    className="w-full bg-transparent text-right focus:outline-none focus:ring-1 focus:ring-[#141414]"
                                  />
                                </td>
                                <td className="border border-[#141414] p-2">{item.lotNo}</td>
                                {isFirstInGroup ? (
                                  <td 
                                    rowSpan={groupSize} 
                                    className="border border-[#141414] p-2 text-right font-mono align-middle bg-gray-50/50"
                                  >
                                    {item.actualIssuedQty?.toLocaleString()}
                                  </td>
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
                                <td className="border border-[#141414] p-2">{item.brand}</td>
                                <td className="border border-[#141414] p-2">{item.customerCode}</td>
                                <td className="border border-[#141414] p-2">{item.finalDestination}</td>
                                <td className="border border-[#141414] p-2">{item.noCode}</td>
                                <td className="border border-[#141414] p-2">{item.location}</td>
                                <td className={cn(
                                  "border border-[#141414] p-2",
                                  item.stock === 'Không có tồn' ? "bg-red-500 text-white font-bold" : ""
                                )}>
                                  {item.stock}
                                </td>
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
                <div className="flex justify-between items-center no-print">
                  <div className="flex items-center gap-6">
                    <h2 className="font-serif italic text-2xl">Quản lý Vị Trí</h2>
                    <div className="flex border-b border-gray-200">
                      <button
                        onClick={() => setLocationSubTab('input')}
                        className={cn(
                          "px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors border-b-2",
                          locationSubTab === 'input' ? "border-[#141414] text-[#141414]" : "border-transparent text-gray-400 hover:text-gray-600"
                        )}
                      >
                        NHẬP VỊ TRÍ
                      </button>
                      <button
                        onClick={() => setLocationSubTab('output')}
                        className={cn(
                          "px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors border-b-2",
                          locationSubTab === 'output' ? "border-[#141414] text-[#141414]" : "border-transparent text-gray-400 hover:text-gray-600"
                        )}
                      >
                        XUẤT VỊ TRÍ
                      </button>
                      <button
                        onClick={() => setLocationSubTab('inventory')}
                        className={cn(
                          "px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors border-b-2",
                          locationSubTab === 'inventory' ? "border-[#141414] text-[#141414]" : "border-transparent text-gray-400 hover:text-gray-600"
                        )}
                      >
                        TỒN VỊ TRÍ
                      </button>
                    </div>
                    <div className="relative ml-4">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={14} />
                      <input 
                        type="text" 
                        placeholder="Tìm kiếm vị trí, mã hàng..."
                        value={locationSearch}
                        onChange={(e) => setLocationSearch(e.target.value)}
                        className="pl-9 pr-4 py-2 bg-white border border-[#141414] text-xs focus:outline-none focus:ring-1 focus:ring-[#141414] w-64"
                      />
                    </div>
                  </div>
                  <div className="flex gap-3">
                    <button 
                      onClick={() => {
                        setEditingId(null);
                        setNewLocationEntry({ qrcode: '', sku: '', partner: '', date: format(new Date(), 'dd/MM/yyyy'), location: '', note: '', quantity: 1 });
                        if (locationSubTab === 'output') {
                          setScanMode('OUTPUT');
                        } else {
                          setScanMode('INPUT');
                        }
                        setIsLocationModalOpen(true);
                      }}
                      className="flex items-center gap-2 px-6 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <Plus size={14} />
                      Thêm vị trí
                    </button>
                    <button 
                      onClick={() => fileInputRef.current?.click()}
                      className="flex items-center gap-2 px-6 py-2 bg-[#141414] text-[#E4E3E0] text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                    >
                      <FileUp size={14} />
                      Nhập Excel {locationSubTab === 'inventory' ? '(Sheet 3)' : ''}
                    </button>
                    {selectedRows.length > 0 && (
                      <button 
                        onClick={handleBulkDelete}
                        className="flex items-center gap-2 px-6 py-2 bg-red-600 text-white text-xs font-bold uppercase tracking-wider hover:opacity-90 transition-opacity"
                      >
                        <Trash2 size={14} />
                        Xóa ({selectedRows.length})
                      </button>
                    )}
                  </div>
                </div>

                {(locationSubTab === 'input' || locationSubTab === 'output') && (
                  <div className="space-y-4">
                    <div className="border border-[#141414] overflow-x-auto">
                    <table className="w-full border-collapse">
                      <thead>
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
                          <th className="border border-[#141414] p-3 text-left">Cuộn</th>
                          <th className="border border-[#141414] p-3 text-left">Vị trí</th>
                          <th className="border border-[#141414] p-3 text-right">SL</th>
                          <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                            Thao tác
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredLocationEntries.length === 0 ? (
                          <tr>
                            <td colSpan={8} className="border border-[#141414] p-8 text-center italic text-gray-400">
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
                                    <td colSpan={7} className="border border-[#141414] p-2 uppercase">
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
                                  <td className="border border-[#141414] p-3 italic">{entry.note}</td>
                                  <td className="border border-[#141414] p-3 font-bold text-blue-600">{entry.location}</td>
                                  <td className="border border-[#141414] p-3 text-right font-bold">{entry.quantity}</td>
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
                  <div className="border border-[#141414] overflow-x-auto">
                    <table className="w-full border-collapse">
                      <thead>
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
                          <th className="border border-[#141414] p-3 text-left">Cuộn</th>
                          <th className="border border-[#141414] p-3 text-left">Vị trí</th>
                          <th className="border border-[#141414] p-3 text-right">SL</th>
                          <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                            Thao tác
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredLocationInventoryEntries.length === 0 ? (
                          <tr>
                            <td colSpan={8} className="border border-[#141414] p-8 text-center italic text-gray-400">
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
                              <td className="border border-[#141414] p-3 italic">{entry.note}</td>
                              <td className="border border-[#141414] p-3 font-bold text-blue-600">{entry.location}</td>
                              <td className="border border-[#141414] p-3 text-right font-bold">{entry.quantity}</td>
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
        </div>
      </main>

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
                      const isSimpleAWB = qrcode.toUpperCase().startsWith('AWB-') && !qrcode.includes('|');
                      setNewLocationEntry({
                        ...newLocationEntry,
                        qrcode,
                        sku: parsed?.sku || newLocationEntry.sku,
                        partner: parsed?.partner || newLocationEntry.partner,
                        date: isSimpleAWB ? '' : (parsed?.date || newLocationEntry.date)
                      });
                    }}
                    className="w-full bg-transparent border-b border-[#141414] py-2 focus:outline-none focus:border-blue-500 transition-colors font-mono"
                    placeholder="Nhập hoặc dán mã QR..."
                  />
                </div>
                <div className="grid grid-cols-2 gap-4">
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
                <div className="grid grid-cols-2 gap-4">
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
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Số lượng</label>
                    <input 
                      type="number" 
                      required
                      min="1"
                      value={newLocationEntry.quantity} 
                      onChange={(e) => setNewLocationEntry({...newLocationEntry, quantity: parseInt(e.target.value) || 0})}
                      className="w-full bg-transparent border-b border-[#141414] py-2 focus:outline-none focus:border-blue-500 transition-colors font-bold"
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] uppercase font-bold opacity-50">Ghi chú</label>
                  <input 
                    type="text" 
                    value={newLocationEntry.note} 
                    onChange={(e) => setNewLocationEntry({...newLocationEntry, note: e.target.value})}
                    className="w-full bg-transparent border-b border-[#141414] py-2 focus:outline-none focus:border-blue-500 transition-colors italic"
                  />
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
                <div className="grid grid-cols-2 gap-4">
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
                <div className="grid grid-cols-2 gap-4">
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
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-[10px] uppercase font-bold opacity-50">Mã chỉ định</label>
                    <input 
                      className="w-full bg-transparent border-b border-[#141414] py-1 text-sm outline-none"
                      value={newProduct.designationCode}
                      onChange={e => setNewProduct({...newProduct, designationCode: e.target.value})}
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
                <div className="grid grid-cols-2 gap-4">
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
                    onClick={() => setIsTransactionModalOpen(false)}
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
                <div className="grid grid-cols-2 gap-4">
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
                <div className="grid grid-cols-3 gap-4">
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
                <div className="grid grid-cols-2 gap-4">
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
                  Bạn có chắc chắn muốn xóa {deleteTarget?.id === 'bulk' ? `hàng loạt (${selectedRows.length} mục)` : 'mục này'} không? Hành động này không thể hoàn tác.
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
