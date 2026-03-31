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
  Save
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

import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

type Tab = 'dashboard' | 'inbound' | 'outbound' | 'inventory' | 'customers' | 'deliveryNote' | 'location';

export default function App() {
  const [activeTab, setActiveTab] = useState<Tab>('dashboard');
  const [products, setProducts] = useState<Product[]>(INITIAL_PRODUCTS);
  const [transactions, setTransactions] = useState<Transaction[]>(INITIAL_TRANSACTIONS);
  const [customers, setCustomers] = useState<Customer[]>(INITIAL_CUSTOMERS);
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
  const [locationSubTab, setLocationSubTab] = useState<'input' | 'inventory'>('input');
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
    qrcode: '', sku: '', partner: '', date: format(new Date(), 'dd/MM/yyyy'), location: '', note: ''
  });
  const [deleteTarget, setDeleteTarget] = useState<{ id: string | 'bulk', type: 'product' | 'transaction' | 'customer' | 'location' | 'savedDeliveryNote' } | null>(null);
  const [isSupabaseConfigured, setIsSupabaseConfigured] = useState(true);
  const [isSaving, setIsSaving] = useState(false);

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

  // Load data from Supabase on start
  React.useEffect(() => {
    const getValidUrl = (url: any): string => {
      const placeholder = 'https://placeholder-project.supabase.co';
      if (typeof url !== 'string') return placeholder;
      try {
        const parsed = new URL(url);
        if (parsed.protocol === 'http:' || parsed.protocol === 'https:') {
          return url;
        }
      } catch {
        // Not a valid URL
      }
      return placeholder;
    };

    const supabaseUrl = getValidUrl(import.meta.env.VITE_SUPABASE_URL);
    const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

    if (!import.meta.env.VITE_SUPABASE_URL || !supabaseAnonKey || supabaseUrl.includes('placeholder')) {
      setIsSupabaseConfigured(false);
    }

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

        // If database is empty, seed it with initial data
        if (dbProducts.length === 0 && INITIAL_PRODUCTS.length > 0) {
          await api.products.upsertAll(INITIAL_PRODUCTS);
          setProducts(INITIAL_PRODUCTS);
        } else if (dbProducts.length > 0) {
          setProducts(dbProducts);
        }

        if (dbTransactions.length === 0 && INITIAL_TRANSACTIONS.length > 0) {
          await api.transactions.upsertAll(INITIAL_TRANSACTIONS);
          setTransactions(INITIAL_TRANSACTIONS);
        } else if (dbTransactions.length > 0) {
          setTransactions(dbTransactions);
        }

        if (dbCustomers.length === 0 && INITIAL_CUSTOMERS.length > 0) {
          await api.customers.upsertAll(INITIAL_CUSTOMERS);
          setCustomers(INITIAL_CUSTOMERS);
        } else if (dbCustomers.length > 0) {
          setCustomers(dbCustomers);
        }

        if (dbDeliveryNotes.length > 0) setDeliveryNotes(dbDeliveryNotes);
        if (dbLocationEntries.length > 0) {
          setLocationEntries(dbLocationEntries.filter(e => e.type === 'input' || !e.type));
          const inventoryEntries = dbLocationEntries.filter(e => e.type === 'inventory');
          // Deduplicate existing inventory entries
          const grouped = new Map<string, LocationEntry>();
          inventoryEntries.forEach(entry => {
            const key = `${entry.qrcode}|${entry.location}`;
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
      } catch (error) {
        console.error('Error loading data from Supabase:', error);
      }
    };

    loadData();
  }, []);

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
      const updatedNotes = deliveryNotes.filter(item => item.id !== deliveryNoteDeleteId);
      setDeliveryNotes(updatedNotes);
      try {
        await api.deliveryNotes.upsertAll(updatedNotes);
      } catch (error) {
        console.error('Error syncing delivery notes:', error);
      }
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
    // This one is still using index because it's called from filteredDeliveryNotes.map
    // I should probably change this too to use ID for safety.
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
      } catch (error) {
        console.error('Error syncing post delivery note:', error);
      }
    } else {
      console.warn('Không có dữ liệu thực tế để xuất kho!');
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
    } catch (error) {
      console.error('Error syncing product:', error);
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
    } catch (error) {
      console.error('Error syncing transaction:', error);
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
    } catch (error) {
      console.error('Error syncing customer:', error);
    }
    
    setIsCustomerModalOpen(false);
    setNewCustomer({ code: '', name: '' });
  };

  const handleLocationSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    let updatedEntry: LocationEntry;
    if (editingId) {
      updatedEntry = { ...newLocationEntry, id: editingId, type: locationSubTab } as LocationEntry;
      if (locationSubTab === 'input') {
        setLocationEntries(locationEntries.map(entry => entry.id === editingId ? updatedEntry : entry));
      } else {
        setLocationInventoryEntries(locationInventoryEntries.map(entry => entry.id === editingId ? updatedEntry : entry));
      }
      setEditingId(null);
    } else {
      // Merge if duplicate QRCODE and Location in inventory
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
          };
          setLocationInventoryEntries(prev => prev.map((e, i) => i === existingIndex ? updatedEntry : e));
          setEditingId(null);
        } else {
          updatedEntry = {
            ...newLocationEntry as LocationEntry,
            id: generateId(),
            type: locationSubTab
          };
          setLocationInventoryEntries(prev => [...prev, updatedEntry]);
        }
      } else {
        updatedEntry = {
          ...newLocationEntry as LocationEntry,
          id: generateId(),
          type: locationSubTab
        };
        setLocationEntries([...locationEntries, updatedEntry]);
      }
    }
    
    try {
      await api.locationEntries.upsert(updatedEntry);
    } catch (error) {
      console.error('Error syncing location entry:', error);
    }
    
    setIsLocationModalOpen(false);
    setNewLocationEntry({ qrcode: '', sku: '', partner: '', date: format(new Date(), 'dd/MM/yyyy'), location: '', note: '' });
  };

  const processLocationData = async (data: any[]) => {
    const entries: any[] = data.map((row: any) => {
      // Normalize row keys
      const normalizedRow: any = {};
      Object.keys(row).forEach(key => {
        normalizedRow[key.toLowerCase().trim()] = row[key];
      });

      const qrcode = normalizedRow['qrcode'] || normalizedRow['qr code'] || row['QR Code'] || row['qrcode'] || '';
      const parsed = parseQRCode(qrcode);

      return {
        qrcode,
        sku: normalizedRow['sku'] || row['SKU'] || row['sku'] || parsed?.sku || '',
        partner: normalizedRow['partner'] || normalizedRow['đối tác'] || row['Đối tác'] || row['partner'] || parsed?.partner || '',
        date: normalizedRow['date'] || normalizedRow['ngày'] || row['Ngày'] || row['date'] || parsed?.date || format(new Date(), 'dd/MM/yyyy'),
        location: normalizedRow['location'] || normalizedRow['vị trí'] || row['Vị trí'] || row['location'] || '',
        note: normalizedRow['note'] || normalizedRow['ghi chú'] || row['Ghi chú'] || row['note'] || ''
      };
    });

    // Group by QR code and merge locations
    const grouped = new Map<string, any>();
    
    // First, add existing entries to the map
    locationEntries.forEach(entry => {
      grouped.set(entry.qrcode, { ...entry });
    });

    // Then, merge new entries
    entries.forEach(entry => {
      if (grouped.has(entry.qrcode)) {
        const existing = grouped.get(entry.qrcode);
        if (entry.location && !existing.location.includes(entry.location)) {
          existing.location = existing.location ? `${existing.location}, ${entry.location}` : entry.location;
        }
        // Update other fields if they were empty
        if (!existing.sku) existing.sku = entry.sku;
        if (!existing.partner) existing.partner = entry.partner;
        if (!existing.date) existing.date = entry.date;
        if (!existing.note) existing.note = entry.note;
      } else {
        grouped.set(entry.qrcode, { ...entry, id: generateId(), type: 'input' });
      }
    });

    const finalEntries: LocationEntry[] = Array.from(grouped.values());
    setLocationEntries(finalEntries);
    
    try {
      await api.locationEntries.upsertAll(finalEntries);
    } catch (error) {
      console.error('Error syncing location entries:', error);
    }
  };

  const handleDelete = async (id: string, type: 'product' | 'transaction' | 'customer' | 'location' | 'savedDeliveryNote') => {
    try {
      if (type === 'product') {
        await api.products.delete(id);
        setProducts(products.filter(p => p.id !== id));
      }
      if (type === 'transaction') {
        await api.transactions.delete(id);
        setTransactions(transactions.filter(t => t.id !== id));
      }
      if (type === 'customer') {
        await api.customers.delete(id);
        setCustomers(customers.filter(c => c.id !== id));
      }
      if (type === 'location') {
        await api.locationEntries.delete(id);
        setLocationEntries(locationEntries.filter(e => e.id !== id));
        setLocationInventoryEntries(locationInventoryEntries.filter(e => e.id !== id));
      }
      if (type === 'savedDeliveryNote') {
        await api.savedDeliveryNotes.delete(id);
        setSavedDeliveryNotes(savedDeliveryNotes.filter(n => n.id !== id));
      }
    } catch (error) {
      console.error('Error syncing deletion:', error);
    }
  };

  const handleBulkDelete = async () => {
    try {
      if (activeTab === 'inventory') {
        await Promise.all(selectedRows.map(id => api.products.delete(id)));
        setProducts(products.filter(p => !selectedRows.includes(p.id)));
      }
      if (activeTab === 'inbound' || activeTab === 'outbound') {
        await Promise.all(selectedRows.map(id => api.transactions.delete(id)));
        setTransactions(transactions.filter(t => !selectedRows.includes(t.id)));
      }
      if (activeTab === 'customers') {
        await Promise.all(selectedRows.map(id => api.customers.delete(id)));
        setCustomers(customers.filter(c => !selectedRows.includes(c.id)));
      }
    } catch (error) {
      console.error('Error syncing bulk deletion:', error);
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
      // Using a simple alert for now as requested
      alert('Dữ liệu đã được lưu trữ thành công!');
    } catch (error) {
      console.error('Lỗi khi lưu dữ liệu:', error);
      alert('Có lỗi xảy ra khi lưu dữ liệu. Vui lòng kiểm tra lại kết nối.');
    } finally {
      setIsSaving(false);
    }
  };

  const confirmDelete = () => {
    if (!deleteTarget) return;
    if (deleteTarget.id === 'bulk') {
      handleBulkDelete();
    } else {
      handleDelete(deleteTarget.id, deleteTarget.type);
    }
    setIsDeleteConfirmOpen(false);
    setDeleteTarget(null);
  };

  const toggleRowSelection = (id: string) => {
    setSelectedRows(prev => prev.includes(id) ? prev.filter(rowId => rowId !== id) : [...prev, id]);
  };

  // Derived Inventory State
  const inventory = useMemo(() => {
    const batches: Record<string, InventoryItem> = {};
    const productMap = new Map<string, Product>(products.map(p => [p.id, p]));
    const productsWithTransactions = new Set<string>();

    // Process all transactions to group by batch
    transactions.forEach(t => {
      const product = productMap.get(t.productId);
      if (!product) return;

      productsWithTransactions.add(t.productId);

      // Define a batch by product ID and its metadata
      const batchKey = `${t.productId}-${t.lotNo || ''}-${t.designationCode || ''}-${t.loaiChiDinh || ''}`;
      
      if (!batches[batchKey]) {
        batches[batchKey] = {
          ...product,
          id: batchKey, // Unique ID for the row in the table
          productId: product.id, // Keep reference to original product
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

    // Also include products that have no transactions yet
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
    return locationEntries.filter(entry => 
      entry.sku.toLowerCase().includes(query) ||
      entry.partner.toLowerCase().includes(query) ||
      entry.location.toLowerCase().includes(query) ||
      entry.qrcode.toLowerCase().includes(query) ||
      entry.note.toLowerCase().includes(query)
    );
  }, [locationEntries, locationSearch]);

  const filteredLocationInventoryEntries = useMemo(() => {
    const query = locationSearch.toLowerCase();
    return locationInventoryEntries.filter(entry => 
      entry.sku.toLowerCase().includes(query) ||
      entry.partner.toLowerCase().includes(query) ||
      entry.location.toLowerCase().includes(query) ||
      entry.qrcode.toLowerCase().includes(query) ||
      entry.note.toLowerCase().includes(query)
    );
  }, [locationInventoryEntries, locationSearch]);

  const parseDate = useCallback((dateStr: string) => {
    // Try dd/MM/yyyy first
    let parsed = parse(dateStr, 'dd/MM/yyyy', new Date());
    if (isValid(parsed)) return parsed;
    
    // Fallback to standard Date parsing
    parsed = new Date(dateStr);
    if (isValid(parsed)) return parsed;

    return new Date(0); // Epoch as fallback
  }, []);

  const filteredTransactions = useMemo(() => {
    const productMap = new Map<string, Product>(products.map(p => [p.id, p]));
    const query = searchQuery.toLowerCase();
    
    // Pre-parse dates for sorting to avoid repeated parsing
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
      // Import/Update Products
      setProducts(prev => {
        const updatedProducts = [...prev];
        const skuToProductIndex = new Map(updatedProducts.map((p, i) => [p.sku, i]));

        data.forEach((row, index) => {
          // Normalize row keys
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
      // Import Transactions
      // First, ensure all products mentioned in the file exist and are up to date
      setProducts(prevProducts => {
        const updatedProducts = [...prevProducts];
        const newProductsToAdd: Product[] = [];
        
        // Optimization: Use maps for faster SKU lookups
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
            // Update name if provided and different
            if (name && updatedProducts[existingIndex].name !== name && name !== 'Sản phẩm mới') {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], name };
            }
            // Update designationCode if provided
            if (designationCode && updatedProducts[existingIndex].designationCode !== designationCode) {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], designationCode };
            }
            // Update loaiChiDinh if provided
            if (loaiChiDinh && updatedProducts[existingIndex].loaiChiDinh !== loaiChiDinh) {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], loaiChiDinh };
            }
            // Update ghiChu if provided
            if (ghiChu && updatedProducts[existingIndex].ghiChu !== ghiChu) {
              updatedProducts[existingIndex] = { ...updatedProducts[existingIndex], ghiChu };
            }
          } else if (alreadyInNewIndex !== undefined) {
            // Update name in the new products list if provided
            if (name && newProductsToAdd[alreadyInNewIndex].name === 'Sản phẩm mới' && name !== 'Sản phẩm mới') {
              newProductsToAdd[alreadyInNewIndex] = { ...newProductsToAdd[alreadyInNewIndex], name };
            }
            // Update designationCode if provided
            if (designationCode && newProductsToAdd[alreadyInNewIndex].designationCode === '') {
              newProductsToAdd[alreadyInNewIndex] = { ...newProductsToAdd[alreadyInNewIndex], designationCode };
            }
            // Update loaiChiDinh if provided
            if (loaiChiDinh && (!newProductsToAdd[alreadyInNewIndex].loaiChiDinh || newProductsToAdd[alreadyInNewIndex].loaiChiDinh === '')) {
              newProductsToAdd[alreadyInNewIndex] = { ...newProductsToAdd[alreadyInNewIndex], loaiChiDinh };
            }
            // Update ghiChu if provided
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

        // Now process transactions using the final products list
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
      // Import Customers
      setCustomers(prev => {
        const updatedCustomers = [...prev];
        const newCustomers: Customer[] = [];
        const codeToCustomerIndex = new Map(updatedCustomers.map((c, i) => [c.code, i]));

        data.forEach((row, index) => {
          const code = String(row.code || row.Code || row['Mã'] || '').trim();
          if (!code) return;

          const existingIndex = codeToCustomerIndex.get(code);
          const customerData: Customer = {
            id: existingIndex !== undefined ? updatedCustomers[existingIndex].id : generateId(),
            code: code,
            name: row.name || row.Name || row['Tên'] || 'Khách hàng mới',
          };

          if (existingIndex !== undefined) {
            updatedCustomers[existingIndex] = customerData;
          } else {
            codeToCustomerIndex.set(code, updatedCustomers.length);
            updatedCustomers.push(customerData);
            newCustomers.push(customerData);
          }
        });
        
        Promise.all(updatedCustomers.map(c => api.customers.upsert(c))).catch(err => console.error('Error syncing customers:', err));
        return updatedCustomers;
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
            location: String(row[4] || '').trim(),
            note: String(row[5] || '').trim(),
          });
        }
      } else if (currentSection === 'OUTPUT') {
        const qrcode = String(row[0] || '').trim();
        if (qrcode && qrcode !== 'QRCODE') {
          outputQRCodes.add(qrcode);
        }
      }
    });

    // Filter out outputted QR codes
    const filtered = inputEntries.filter(entry => !outputQRCodes.has(entry.qrcode));
    
    // Group by QR code and location to prevent duplicates
    const grouped = new Map<string, any>();
    filtered.forEach(entry => {
      const key = `${entry.qrcode}|${entry.location}`;
      if (!grouped.has(key)) {
        grouped.set(key, { ...entry });
      } else {
        // If duplicate (same QR and same Location), we can update other fields if they were empty
        const existing = grouped.get(key);
        if (!existing.sku) existing.sku = entry.sku;
        if (!existing.partner) existing.partner = entry.partner;
        if (!existing.date) existing.date = entry.date;
        if (!existing.note) existing.note = entry.note;
      }
    });

    return Array.from(grouped.values());
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
          // Use 3rd sheet for location inventory (index 2)
          const sheetName = workbook.SheetNames[2] || workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
          const processedData = processLocationStockSheet(sheetData);
          
          const newEntries: LocationEntry[] = processedData.map((item, index) => ({
            id: generateId(),
            ...item,
            type: 'inventory'
          }));
          
          // Merge with existing inventory entries
          const grouped = new Map<string, LocationEntry>();
          locationInventoryEntries.forEach(entry => {
            const key = `${entry.qrcode}|${entry.location}`;
            grouped.set(key, { ...entry });
          });
          
          newEntries.forEach(entry => {
            const key = `${entry.qrcode}|${entry.location}`;
            if (grouped.has(key)) {
              const existing = grouped.get(key)!;
              // Update fields if they were empty in the existing entry
              if (!existing.sku) existing.sku = entry.sku;
              if (!existing.partner) existing.partner = entry.partner;
              if (!existing.date) existing.date = entry.date;
              if (!existing.note) existing.note = entry.note;
            } else {
              grouped.set(key, entry);
            }
          });
          
          const finalEntries = Array.from(grouped.values());
          setLocationInventoryEntries(finalEntries);
          api.locationEntries.upsertAll(finalEntries).catch(err => console.error('Error syncing location inventory:', err));
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
    
    // Group inventory by SKU for faster lookup
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
      // Try slash format: dd/mm/yy or dd/mm/yyyy
      const slashMatch = lot.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
      if (slashMatch) {
        let [_, d, m, y] = slashMatch;
        const year = y.length === 2 ? 2000 + parseInt(y) : parseInt(y);
        return new Date(year, parseInt(m) - 1, parseInt(d)).getTime();
      }
      
      // Try hyphen format: dd-mm-yyyy or dd-mm-yy
      const hyphenMatch = lot.match(/(\d{1,2})-(\d{1,2})-(\d{2,4})/);
      if (hyphenMatch) {
        let [_, d, m, y] = hyphenMatch;
        const year = y.length === 2 ? 2000 + parseInt(y) : parseInt(y);
        return new Date(year, parseInt(m) - 1, parseInt(d)).getTime();
      }

      // Try yyyy-mm-dd or yy-mm-dd
      const isoMatch = lot.match(/(\d{2,4})-(\d{1,2})-(\d{1,2})/);
      if (isoMatch) {
        let [_, y, m, d] = isoMatch;
        const year = y.length === 2 ? 2000 + parseInt(y) : parseInt(y);
        return new Date(year, parseInt(m) - 1, parseInt(d)).getTime();
      }

      // If no date found, return a very large number so it's picked last in FIFO
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

    // Track remaining stock during allocation to respect FIFO across multiple rows
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

      // Find best stock based on priority and FIFO
      const availableStock = workingInventory.filter(item => item.sku === itemNo && item.tempStock > 0);
      
      let bestBatch: typeof workingInventory[0] | undefined;
      
      for (const priorityType of priorityOrder) {
        const matchingBatches = availableStock.filter(item => 
          (item.loaiChiDinh || '').trim().toUpperCase() === priorityType.toUpperCase()
        );
        
        if (matchingBatches.length > 0) {
          // Sort by date (FIFO) within the same priority
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
            // For other types, take the oldest available
            bestBatch = matchingBatches[0];
            break;
          }
        }
      }

      if (bestBatch) {
        bestBatch.tempStock -= qtyNeeded;
      }

      // Look up location from locationInventoryEntries
      const normalizedLotDate = normalizeDateForMatching(bestBatch?.lotNo || '');
      const locationEntry = locationInventoryEntries.find(e => 
        e.sku === itemNo && e.date === normalizedLotDate
      );

      return {
        id: generateId(),
        no: 0, // Will be re-indexed
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

    // Sort by item (A-Z)
    const sortedData = mappedData.sort((a, b) => a.item.localeCompare(b.item));
    
    // Group and calculate adjustments
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
        
        // Find row with largest Qty ERP among valid rows
        let maxRow = validRows[0];
        validRows.forEach(item => {
          if (item.qtyErp > maxRow.qtyErp) maxRow = item;
        });
        
        // Adjust max row
        maxRow.actualQty = maxRow.qtyErp + diff;
        
        // Set actualIssuedQty for all rows in group (total rounded of valid rows)
        group.forEach(item => {
          item.actualIssuedQty = roundedTotal;
        });
      } else {
        // All rows in group are out of stock
        group.forEach(item => {
          item.actualQty = 0;
          item.actualIssuedQty = 0;
        });
      }
    });

    // Re-index No.
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
        // Create a normalized version of the row with lowercase keys for easier matching
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
          // Check if already in existing customers
          const existsInCurrent = currentCustomers.find(c => 
            c.code.toLowerCase() === code.toLowerCase() && 
            c.name.toLowerCase() === name.toLowerCase()
          );
          // Check if already in newCustomers (to avoid duplicates within the file)
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
      {/* Sidebar */}
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

      {/* Main Content */}
      <main className="flex-1 flex flex-col overflow-hidden">
        {/* Header */}
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
              className="flex items-center gap-2 px-4 py-2 bg-[#141414] text-[#E4E3E0] text-[10px] font-bold uppercase tracking-wider hover:opacity-90 transition-opacity disabled:opacity-50"
            >
              {isSaving ? (
                <div className="w-3 h-3 border-2 border-[#E4E3E0] border-t-transparent rounded-full animate-spin" />
              ) : (
                <Save size={14} />
              )}
              {isSaving ? 'Đang lưu...' : 'Lưu tất cả'}
            </button>
            {selectedRows.length > 0 && (
              <button 
                onClick={() => {
                  setDeleteTarget({ id: 'bulk', type: activeTab === 'inventory' ? 'product' : (activeTab === 'customers' ? 'customer' : 'transaction') });
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

        {/* Supabase Warning Banner */}
        {!isSupabaseConfigured && (
          <div className="bg-amber-100 border-b border-amber-200 px-8 py-2 flex items-center justify-between">
            <div className="flex items-center gap-3 text-amber-800 text-[10px]">
              <AlertTriangle size={14} />
              <span>
                <strong>Supabase chưa được cấu hình:</strong> Vui lòng thiết lập <code>VITE_SUPABASE_URL</code> và <code>VITE_SUPABASE_ANON_KEY</code> trong menu Settings để đồng bộ dữ liệu.
              </span>
            </div>
          </div>
        )}

        {/* Content Area */}
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
                      <ResponsiveContainer width="100%" height="100%">
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

                        // Set column widths
                        worksheet.columns = [
                          { width: 5 },  // No
                          { width: 20 }, // OVN Sale Order
                          { width: 20 }, // OVN Production Order
                          { width: 15 }, // item
                          { width: 35 }, // Material Name
                          { width: 8 },  // Unit
                          { width: 12 }, // Qty ERP
                          { width: 12 }, // Thực tế
                          { width: 15 }, // Lot No
                          { width: 15 }, // Số lượng thực phát
                          { width: 15 }, // remark
                          { width: 15 }, // Brand
                          { width: 20 }, // Customer code
                          { width: 20 }, // Final Destination
                          { width: 10 }, // No.
                          { width: 10 }, // Vị trí
                          { width: 10 }  // STOCK
                        ];

                        // Title
                        const titleRow = worksheet.addRow(['', 'PHIẾU GIAO NHẬN FABRIC']);
                        titleRow.getCell(2).font = { name: 'Times New Roman', size: 16, bold: true };
                        titleRow.getCell(2).alignment = { horizontal: 'center' };
                        worksheet.mergeCells(1, 2, 1, 17);

                        const subtitleRow = worksheet.addRow(['', 'Delivery Note']);
                        subtitleRow.getCell(2).font = { name: 'Times New Roman', size: 12, italic: true };
                        subtitleRow.getCell(2).alignment = { horizontal: 'center' };
                        worksheet.mergeCells(2, 2, 2, 17);

                        worksheet.addRow([]); // Empty row

                        // Metadata
                        const meta1 = worksheet.addRow(['Mã Tài Liệu:', deliveryNoteHeader.documentCode]);
                        meta1.getCell(1).font = { bold: true };
                        const meta2 = worksheet.addRow(['Dept:', deliveryNoteHeader.dept]);
                        meta2.getCell(1).font = { bold: true };
                        const meta3 = worksheet.addRow(['TO:', deliveryNoteHeader.to]);
                        meta3.getCell(1).font = { bold: true };
                        const meta4 = worksheet.addRow(['Date:', deliveryNoteHeader.date]);
                        meta4.getCell(1).font = { bold: true };

                        worksheet.addRow([]); // Empty row before table

                        // Header
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

                        // Data
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
                            
                            // Alignments
                            if ([1, 6, 10].includes(colNumber)) {
                              cell.alignment = { horizontal: 'center' };
                            } else if ([7, 8].includes(colNumber)) {
                              cell.alignment = { horizontal: 'right' };
                            }
                          });
                        });

                        // Merges for "Số lượng thực phát" (Column 10)
                        let currentItem = '';
                        let startMergeRow = 0;
                        const dataStartRow = 9; // Header is at row 8, data starts at 9

                        deliveryNotes.forEach((item, index) => {
                          const rowIdx = dataStartRow + index;
                          if (item.item !== currentItem) {
                            if (startMergeRow !== 0 && (rowIdx - 1) > startMergeRow) {
                              worksheet.mergeCells(startMergeRow, 10, rowIdx - 1, 10);
                            }
                            currentItem = item.item;
                            startMergeRow = rowIdx;
                          }
                          // Last group
                          if (index === deliveryNotes.length - 1) {
                            if (rowIdx > startMergeRow) {
                              worksheet.mergeCells(startMergeRow, 10, rowIdx, 10);
                            }
                          }
                        });

                        // Add Total Row
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

                        // Footer / Signatures
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

                        // Center all signature cells
                        [signHeader, signSub].forEach(row => {
                          row.eachCell(cell => { cell.alignment = { horizontal: 'center' }; });
                        });

                        // Generate and save
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
                  {/* Header Template */}
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
                        setNewLocationEntry({ qrcode: '', sku: '', partner: '', date: format(new Date(), 'dd/MM/yyyy'), location: '', note: '' });
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
                  </div>
                </div>

                {locationSubTab === 'input' && (
                  <div className="space-y-4">
                    <div className="flex items-center gap-4 bg-yellow-50 p-4 border border-yellow-200">
                      <div className="flex-1">
                        <label className="text-[10px] uppercase font-bold text-yellow-800 block mb-1">QUÉT MÃ VẠCH (VỊ TRÍ HOẶC QRCODE HÀNG)</label>
                        <textarea 
                          rows={3}
                          placeholder="Quét mã vị trí (FB...) hoặc QR code hàng... (Hỗ trợ INPUT/OUTPUT)"
                          className="w-full bg-white border border-yellow-400 py-2 px-3 text-sm focus:outline-none focus:ring-1 focus:ring-yellow-600 font-mono resize-none"
                          onKeyDown={async (e) => {
                            if (e.key === 'Enter' && !e.shiftKey) {
                              e.preventDefault();
                              const text = e.currentTarget.value.trim();
                              if (!text) return;
                              
                              const lines = text.split('\n').map(l => l.trim()).filter(l => l);
                              let tempMode = scanMode;
                              let tempLocation = currentLocation;
                              
                              const newInventoryEntries = [...locationInventoryEntries];
                              const newLogEntries = [...locationEntries];
                              let changed = false;
                              let logChanged = false;

                              for (const line of lines) {
                                const upperLine = line.toUpperCase();
                                if (upperLine === 'INPUT') {
                                  tempMode = 'INPUT';
                                  setScanMode('INPUT');
                                  continue;
                                }
                                if (upperLine === 'OUTPUT') {
                                  tempMode = 'OUTPUT';
                                  setScanMode('OUTPUT');
                                  continue;
                                }
                                
                                if (upperLine.startsWith('FB')) {
                                  tempLocation = line;
                                  setCurrentLocation(line);
                                  continue;
                                }
                                
                                const parsed = parseQRCode(line);
                                if (parsed) {
                                  const logEntry: LocationEntry = {
                                    id: generateId(),
                                    qrcode: line,
                                    sku: parsed.sku,
                                    partner: parsed.partner,
                                    date: parsed.date,
                                    location: tempLocation,
                                    note: '',
                                    type: 'input',
                                    scanType: tempMode
                                  };
                                  newLogEntries.unshift(logEntry);
                                  logChanged = true;
                                  try {
                                    await api.locationEntries.upsert(logEntry);
                                  } catch (error) {
                                    console.error('Error syncing log entry:', error);
                                  }

                                  if (tempMode === 'INPUT') {
                                    // Merge if duplicate in inventory
                                    const existingIndex = newInventoryEntries.findIndex(
                                      entry => entry.qrcode === line && entry.location === tempLocation
                                    );

                                    if (existingIndex >= 0) {
                                      const existing = newInventoryEntries[existingIndex];
                                      const updatedEntry: LocationEntry = {
                                        ...existing,
                                        sku: parsed.sku || existing.sku,
                                        partner: parsed.partner || existing.partner,
                                        date: parsed.date || existing.date,
                                      };
                                      newInventoryEntries[existingIndex] = updatedEntry;
                                      changed = true;
                                      try {
                                        await api.locationEntries.upsert(updatedEntry);
                                      } catch (error) {
                                        console.error('Error syncing location entry:', error);
                                      }
                                    } else {
                                      const newEntry: LocationEntry = {
                                        id: generateId(),
                                        qrcode: line,
                                        sku: parsed.sku,
                                        partner: parsed.partner,
                                        date: parsed.date,
                                        location: tempLocation,
                                        note: '',
                                        type: 'inventory'
                                      };
                                      newInventoryEntries.unshift(newEntry);
                                      changed = true;
                                      try {
                                        await api.locationEntries.upsert(newEntry);
                                      } catch (error) {
                                        console.error('Error syncing location entry:', error);
                                      }
                                    }
                                  } else if (tempMode === 'OUTPUT') {
                                    // Remove from inventory
                                    const existingIndex = newInventoryEntries.findIndex(
                                      entry => entry.qrcode === line && entry.location === tempLocation
                                    );
                                    if (existingIndex >= 0) {
                                      const entryToDelete = newInventoryEntries[existingIndex];
                                      newInventoryEntries.splice(existingIndex, 1);
                                      changed = true;
                                      try {
                                        await api.locationEntries.delete(entryToDelete.id);
                                      } catch (error) {
                                        console.error('Error deleting location entry:', error);
                                      }
                                    }
                                  }
                                }
                              }
                              
                              if (changed) {
                                setLocationInventoryEntries(newInventoryEntries);
                              }
                              if (logChanged) {
                                setLocationEntries(newLogEntries);
                              }
                              e.currentTarget.value = '';
                            }
                          }}
                        />
                      </div>
                      <div className="flex flex-col gap-2">
                        <div className="bg-white border border-yellow-400 p-2 min-w-[150px]">
                          <label className="text-[10px] uppercase font-bold text-gray-500 block">CHẾ ĐỘ QUÉT</label>
                          <div className="flex items-center gap-2">
                            <div className={cn(
                              "w-3 h-3 rounded-full",
                              scanMode === 'INPUT' ? "bg-green-500" : "bg-red-500"
                            )} />
                            <span className={cn(
                              "text-sm font-bold uppercase",
                              scanMode === 'INPUT' ? "text-green-600" : "text-red-600"
                            )}>
                              {scanMode === 'INPUT' ? 'NHẬP VỊ TRÍ' : 'XUẤT VỊ TRÍ'}
                            </span>
                          </div>
                        </div>
                        <div className="bg-white border border-yellow-400 p-2 min-w-[150px]">
                          <label className="text-[10px] uppercase font-bold text-gray-500 block">VỊ TRÍ HIỆN TẠI</label>
                          <span className="text-lg font-bold text-blue-600 font-mono">{currentLocation || 'CHƯA CHỌN'}</span>
                        </div>
                      </div>
                    </div>

                    <div className="border border-[#141414] overflow-x-auto">
                    <table className="w-full border-collapse">
                      <thead>
                        <tr className="bg-[#001F3F] text-white text-[11px] uppercase tracking-wider">
                          <th className="border border-[#141414] p-3 text-left">QRCODE</th>
                          <th className="border border-[#141414] p-3 text-left">Mã</th>
                          <th className="border border-[#141414] p-3 text-left">NCC</th>
                          <th className="border border-[#141414] p-3 text-left">NGÀY</th>
                          <th className="border border-[#141414] p-3 text-left">Cuộn</th>
                          <th className="border border-[#141414] p-3 text-left">Vị trí</th>
                          <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                            Thao tác
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredLocationEntries.length === 0 ? (
                          <tr>
                            <td colSpan={7} className="border border-[#141414] p-8 text-center italic text-gray-400">
                              Chưa có dữ liệu vị trí. Hãy nhập file Excel.
                            </td>
                          </tr>
                        ) : (
                          (() => {
                            const rows: React.ReactNode[] = [];
                            let lastScanType: string | undefined = undefined;
                            
                            // Sort chronologically for the log view to show headers correctly
                            // Assuming new entries are unshifted, we reverse to show chronological order for header logic
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
                                <tr key={entry.id} className="bg-white text-xs hover:bg-gray-50 transition-colors">
                                  <td className="border border-[#141414] p-3 font-mono">{entry.qrcode}</td>
                                  <td className="border border-[#141414] p-3 font-bold">{entry.sku}</td>
                                  <td className="border border-[#141414] p-3">{entry.partner}</td>
                                  <td className="border border-[#141414] p-3">{entry.date}</td>
                                  <td className="border border-[#141414] p-3 italic">{entry.note}</td>
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
                            
                            // Reverse back to show newest at top if that's preferred, 
                            // but headers logic works best chronologically.
                            // Let's keep it newest at top but with headers.
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
                          <th className="border border-[#141414] p-3 text-left">QRCODE</th>
                          <th className="border border-[#141414] p-3 text-left">Mã</th>
                          <th className="border border-[#141414] p-3 text-left">NCC</th>
                          <th className="border border-[#141414] p-3 text-left">NGÀY</th>
                          <th className="border border-[#141414] p-3 text-left">Cuộn</th>
                          <th className="border border-[#141414] p-3 text-left">Vị trí</th>
                          <th className="border border-[#141414] p-3 text-center bg-white text-red-600 font-bold w-32">
                            Thao tác
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredLocationInventoryEntries.length === 0 ? (
                          <tr>
                            <td colSpan={7} className="border border-[#141414] p-8 text-center italic text-gray-400">
                              Không có dữ liệu tồn theo vị trí. Hãy nhập file Excel (Sheet 3).
                            </td>
                          </tr>
                        ) : (
                          filteredLocationInventoryEntries.map(entry => (
                            <tr key={entry.id} className="bg-white text-xs hover:bg-gray-50 transition-colors">
                              <td className="border border-[#141414] p-3 font-mono">{entry.qrcode}</td>
                              <td className="border border-[#141414] p-3 font-bold">{entry.sku}</td>
                              <td className="border border-[#141414] p-3">{entry.partner}</td>
                              <td className="border border-[#141414] p-3">{entry.date}</td>
                              <td className="border border-[#141414] p-3 italic">{entry.note}</td>
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
                          ))
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

      {/* Location Modal */}
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
                        date: parsed?.date || newLocationEntry.date
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

      {/* Product Modal */}
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

      {/* Transaction Modal */}
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

      {/* Customer Modal */}
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

      {/* Delivery Note Item Edit Modal */}
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

      {/* Delivery Note Delete Confirmation Modal */}
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

      {/* Delete Confirmation Modal */}
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
