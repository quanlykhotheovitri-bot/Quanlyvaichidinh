import { supabase } from './supabase';
import { Product, Transaction, Customer, DeliveryNoteItem, LocationEntry } from '../types';

const CACHE_KEY_PREFIX = 'ortholite_cache_';

const getCache = <T>(key: string): T | null => {
  const data = localStorage.getItem(CACHE_KEY_PREFIX + key);
  if (!data) return null;
  try {
    const { value, expiry } = JSON.parse(data);
    if (Date.now() > expiry) {
      localStorage.removeItem(CACHE_KEY_PREFIX + key);
      return null;
    }
    return value as T;
  } catch (e) {
    return null;
  }
};

const setCache = <T>(key: string, value: T, ttl: number = 3600000): void => {
  const expiry = Date.now() + ttl;
  localStorage.setItem(CACHE_KEY_PREFIX + key, JSON.stringify({ value, expiry }));
};

const updateCacheSingle = <T extends {id: string}>(key: string, item: T) => {
  const current = getCache<T[]>(key) || [];
  const index = current.findIndex(p => p.id === item.id);
  if (index >= 0) {
    current[index] = item;
  } else {
    current.push(item);
  }
  setCache(key, current, 30 * 24 * 3600000);
};

const chunkArray = <T>(array: T[], size: number): T[][] => {
  const result: T[][] = [];
  for (let i = 0; i < array.length; i += size) {
    result.push(array.slice(i, i + size));
  }
  return result;
};

const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

export const api = {
  products: {
    async getAll(): Promise<Product[]> {
      try {
        const { data, error } = await supabase.from('products').select('*');
        if (error) throw error;
        
        const mappedData = (data || []).map(row => ({
          ...row,
          minStock: row.minstock ?? row.minStock ?? 0,
          lotNo: row.lotno ?? row.lotNo ?? '',
          ghiChu: row.ghichu ?? row.ghiChu ?? '',
          designationCode: row.designationcode ?? row.designationCode ?? '',
          loaiChiDinh: row.loaichidinh ?? row.loaiChiDinh ?? ''
        })) as Product[];
        setCache('products', mappedData, 30 * 24 * 3600000); // 30 days
        return mappedData;
      } catch (err: any) {
        console.warn('Offline mode: using cached products', err.message);
        const cached = getCache<Product[]>('products');
        if (cached) return cached;
        throw err;
      }
    },
    async upsert(product: Product): Promise<void> {
      updateCacheSingle('products', product);
      const payload = {
        ...product,
        minstock: product.minStock,
        lotno: product.lotNo,
        ghichu: product.ghiChu,
        designationcode: product.designationCode,
        loaichidinh: product.loaiChiDinh
      };
      // Delete camelCase keys to avoid PGRST204
      delete (payload as any).minStock;
      delete (payload as any).lotNo;
      delete (payload as any).ghiChu;
      delete (payload as any).designationCode;
      delete (payload as any).loaiChiDinh;

      try {
        const { error } = await supabase.from('products').upsert(payload);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: product upsert saved locally via full sync');
      }
    },
    async upsertAll(products: Product[]): Promise<void> {
      if (!products.length) return;
      const payload = products.map(product => {
        const p = {
          ...product,
          minstock: product.minStock,
          lotno: product.lotNo,
          ghichu: product.ghiChu,
          designationcode: product.designationCode,
          loaichidinh: product.loaiChiDinh
        };
        delete (p as any).minStock;
        delete (p as any).lotNo;
        delete (p as any).ghiChu;
        delete (p as any).designationCode;
        delete (p as any).loaiChiDinh;
        return p;
      });
      setCache('products', products, 30 * 24 * 3600000); // Optimistic cache update
      try {
        const { error } = await supabase.from('products').upsert(payload);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: upsertAll saved locally');
      }
    },
    async delete(id: string): Promise<void> {
      // 1. Optimistic Cache Update
      const currentProducts = getCache<Product[]>('products') || [];
      setCache('products', currentProducts.filter(p => p.id !== id), 30 * 24 * 3600000);
      
      // 2. Database Update
      const { error } = await supabase.from('products').delete().eq('id', id);
      if (error) {
        console.error('Database error during product deletion:', error);
        throw error;
      }
    },
    async deleteMany(ids: string[]): Promise<void> {
      if (!ids.length) return;
      const currentProducts = getCache<Product[]>('products') || [];
      setCache('products', currentProducts.filter(p => !ids.includes(p.id)), 30 * 24 * 3600000);
      
      const chunks = chunkArray(ids, 100);
      for (const chunk of chunks) {
        const { error } = await supabase.from('products').delete().in('id', chunk);
        if (error) {
          console.error('Database error during products bulk deletion:', error);
          throw error;
        }
        await delay(50);
      }
    },
    async deleteAll(): Promise<void> {
      setCache('products', [], 30 * 24 * 3600000);
      const { error } = await supabase.from('products').delete().neq('id', '');
      if (error) {
        console.error('Database error during products bulk deletion:', error);
        throw error;
      }
    }
  },
  transactions: {
    async getAll(): Promise<Transaction[]> {
      try {
        const { data, error } = await supabase.from('transactions').select('*');
        if (error) throw error;
        
        const mappedData = (data || []).map(row => ({
          ...row,
          productId: row.productid ?? row.productId ?? '',
          loaiChiDinh: row.loaichidinh ?? row.loaiChiDinh ?? '',
          lotNo: row.lotno ?? row.lotNo ?? '',
          ghiChu: row.ghichu ?? row.ghiChu ?? '',
          designationCode: row.designationcode ?? row.designationCode ?? ''
        })) as Transaction[];
        setCache('transactions', mappedData, 30 * 24 * 3600000);
        return mappedData;
      } catch (err: any) {
        console.warn('Offline mode: using cached transactions', err.message);
        const cached = getCache<Transaction[]>('transactions');
        if (cached) return cached;
        throw err;
      }
    },
    async upsert(transaction: Transaction): Promise<void> {
      updateCacheSingle('transactions', transaction);
      const payload = {
        ...transaction,
        productid: transaction.productId,
        loaichidinh: transaction.loaiChiDinh,
        lotno: transaction.lotNo,
        ghichu: transaction.ghiChu,
        designationcode: transaction.designationCode
      };
      delete (payload as any).productId;
      delete (payload as any).loaiChiDinh;
      delete (payload as any).lotNo;
      delete (payload as any).ghiChu;
      delete (payload as any).designationCode;

      try {
        const { error } = await supabase.from('transactions').upsert(payload);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: transaction upsert saved locally via full sync');
      }
    },
    async upsertAll(transactions: Transaction[]): Promise<void> {
      if (!transactions.length) return;
      const payload = transactions.map(transaction => {
        const t = {
          ...transaction,
          productid: transaction.productId,
          loaichidinh: transaction.loaiChiDinh,
          lotno: transaction.lotNo,
          ghichu: transaction.ghiChu,
          designationcode: transaction.designationCode
        };
        delete (t as any).productId;
        delete (t as any).loaiChiDinh;
        delete (t as any).lotNo;
        delete (t as any).ghiChu;
        delete (t as any).designationCode;
        return t;
      });
      setCache('transactions', transactions, 30 * 24 * 3600000);
      try {
        const { error } = await supabase.from('transactions').upsert(payload);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: upsertAll saved locally');
      }
    },
    async delete(id: string): Promise<void> {
      const currentTransactions = getCache<Transaction[]>('transactions') || [];
      setCache('transactions', currentTransactions.filter(t => t.id !== id), 30 * 24 * 3600000);
      
      const { error } = await supabase.from('transactions').delete().eq('id', id);
      if (error) {
        console.error('Database error during transaction deletion:', error);
        throw error;
      }
    },
    async deleteMany(ids: string[]): Promise<void> {
      if (!ids.length) return;
      const currentTransactions = getCache<Transaction[]>('transactions') || [];
      setCache('transactions', currentTransactions.filter(t => !ids.includes(t.id)), 30 * 24 * 3600000);
      
      const chunks = chunkArray(ids, 100);
      for (const chunk of chunks) {
        const { error } = await supabase.from('transactions').delete().in('id', chunk);
        if (error) {
          console.error('Database error during transactions bulk deletion:', error);
          throw error;
        }
        await delay(50);
      }
    },
    async deleteByProductId(productId: string): Promise<void> {
      // 1. Optimistic Cache Update
      const currentTransactions = getCache<Transaction[]>('transactions') || [];
      setCache('transactions', currentTransactions.filter(t => t.productId !== productId), 30 * 24 * 3600000);
      
      // 2. Database Update
      const { error } = await supabase.from('transactions').delete().eq('productid', productId);
      if (error) {
        console.error('Database error during bulk transaction deletion:', error);
        throw error;
      }
    },
    async deleteByProductIds(productIds: string[]): Promise<void> {
      if (!productIds.length) return;
      const currentTransactions = getCache<Transaction[]>('transactions') || [];
      setCache('transactions', currentTransactions.filter(t => !productIds.includes(t.productId)), 30 * 24 * 3600000);
      
      const chunks = chunkArray(productIds, 100);
      for (const chunk of chunks) {
        const { error } = await supabase.from('transactions').delete().in('productid', chunk);
        if (error) {
          console.error('Database error during bulk transactions deletion by product ids:', error);
          throw error;
        }
        await delay(50);
      }
    },
    async deleteAll(): Promise<void> {
      setCache('transactions', [], 30 * 24 * 3600000);
      const { error } = await supabase.from('transactions').delete().neq('id', '');
      if (error) {
        console.error('Database error during transactions bulk deletion:', error);
        throw error;
      }
    }
  },
  customers: {
    async getAll(): Promise<Customer[]> {
      try {
        const { data, error } = await supabase.from('customers').select('*');
        if (error) throw error;
        setCache('customers', data as Customer[], 30 * 24 * 3600000);
        return data as Customer[];
      } catch (err: any) {
        console.warn('Offline mode: using cached customers', err.message);
        const cached = getCache<Customer[]>('customers');
        if (cached) return cached;
        throw err;
      }
    },
    async upsert(customer: Customer): Promise<void> {
      updateCacheSingle('customers', customer);
      try {
        const { error } = await supabase.from('customers').upsert(customer);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: upsert saved locally');
      }
    },
    async upsertAll(customers: Customer[]): Promise<void> {
      if (!customers.length) return;
      setCache('customers', customers, 30 * 24 * 3600000);
      try {
        const { error } = await supabase.from('customers').upsert(customers);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: upsertAll saved locally');
      }
    },
    async delete(id: string): Promise<void> {
      const currentCustomers = getCache<Customer[]>('customers') || [];
      setCache('customers', currentCustomers.filter(c => c.id !== id), 30 * 24 * 3600000);
      
      const { error } = await supabase.from('customers').delete().eq('id', id);
      if (error) {
        console.error('Database error during customer deletion:', error);
        throw error;
      }
    },
    async deleteMany(ids: string[]): Promise<void> {
      if (!ids.length) return;
      const currentCustomers = getCache<Customer[]>('customers') || [];
      setCache('customers', currentCustomers.filter(c => !ids.includes(c.id)), 30 * 24 * 3600000);
      
      const chunks = chunkArray(ids, 100);
      for (const chunk of chunks) {
        const { error } = await supabase.from('customers').delete().in('id', chunk);
        if (error) {
          console.error('Database error during customers bulk deletion:', error);
          throw error;
        }
        await delay(50);
      }
    }
  },
  deliveryNotes: {
    async getAll(): Promise<DeliveryNoteItem[]> {
      try {
        const { data, error } = await supabase.from('delivery_notes').select('*');
        if (error) throw error;
        
        const mappedData = (data || []).map(row => ({
          ...row,
          ovnSaleOrder: row.ovnsaleorder ?? row.ovnSaleOrder ?? '',
          ovnProductionOrder: row.ovnproductionorder ?? row.ovnProductionOrder ?? '',
          materialName: row.materialname ?? row.materialName ?? '',
          qtyErp: row.qtyerp ?? row.qtyErp ?? 0,
          actualQty: row.actualqty ?? row.actualQty ?? 0,
          lotNo: row.lotno ?? row.lotNo ?? '',
          actualIssuedQty: row.actualissuedqty ?? row.actualIssuedQty ?? 0,
          customerCode: row.customercode ?? row.customerCode ?? '',
          finalDestination: row.finaldestination ?? row.finalDestination ?? '',
          noCode: row.nocode ?? row.noCode ?? ''
        })) as DeliveryNoteItem[];
        setCache('delivery_notes', mappedData, 30 * 24 * 3600000);
        return mappedData;
      } catch (err: any) {
        console.warn('Offline mode: using cached delivery notes', err.message);
        const cached = getCache<DeliveryNoteItem[]>('delivery_notes');
        if (cached) return cached;
        throw err;
      }
    },
    async upsertAll(items: DeliveryNoteItem[]): Promise<void> {
      if (items.length === 0) return;
      const payload = items.map(item => {
        const i = {
          ...item,
          ovnsaleorder: item.ovnSaleOrder,
          ovnproductionorder: item.ovnProductionOrder,
          materialname: item.materialName,
          qtyerp: item.qtyErp,
          actualqty: item.actualQty,
          lotno: item.lotNo,
          actualissuedqty: item.actualIssuedQty,
          customercode: item.customerCode,
          finaldestination: item.finalDestination,
          nocode: item.noCode
        };
        delete (i as any).ovnSaleOrder;
        delete (i as any).ovnProductionOrder;
        delete (i as any).materialName;
        delete (i as any).qtyErp;
        delete (i as any).actualQty;
        delete (i as any).lotNo;
        delete (i as any).actualIssuedQty;
        delete (i as any).customerCode;
        delete (i as any).finalDestination;
        delete (i as any).noCode;
        return i;
      });
      setCache('delivery_notes', items, 30 * 24 * 3600000);
      try {
        const { error } = await supabase.from('delivery_notes').upsert(payload);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: upsertAll delivery notes saved locally');
      }
    },
    async delete(id: string): Promise<void> {
      const currentNotes = getCache<DeliveryNoteItem[]>('delivery_notes') || [];
      setCache('delivery_notes', currentNotes.filter(n => n.id !== id), 30 * 24 * 3600000);
      
      const { error } = await supabase.from('delivery_notes').delete().eq('id', id);
      if (error) {
        console.error('Database error during delivery note deletion:', error);
        throw error;
      }
    },
    async deleteMany(ids: string[]): Promise<void> {
      if (!ids.length) return;
      const currentNotes = getCache<DeliveryNoteItem[]>('delivery_notes') || [];
      setCache('delivery_notes', currentNotes.filter(n => !ids.includes(n.id)), 30 * 24 * 3600000);
      
      const chunks = chunkArray(ids, 100);
      for (const chunk of chunks) {
        const { error } = await supabase.from('delivery_notes').delete().in('id', chunk);
        if (error) {
          console.error('Database error during delivery notes bulk deletion:', error);
          throw error;
        }
        await delay(50);
      }
    },
    async deleteAll(): Promise<void> {
      setCache('delivery_notes', [], 30 * 24 * 3600000);
      const { error } = await supabase.from('delivery_notes').delete().neq('id', '');
      if (error) {
        console.error('Database error during delivery notes deletion:', error);
        throw error;
      }
    }
  },
  locationEntries: {
    async getAll(): Promise<LocationEntry[]> {
      try {
        const { data, error } = await supabase.from('location_entries').select('*');
        if (error) throw error;
        setCache('location_entries', data as LocationEntry[], 30 * 24 * 3600000);
        return data as LocationEntry[];
      } catch (err: any) {
        console.warn('Offline mode: using cached location entries', err.message);
        const cached = getCache<LocationEntry[]>('location_entries');
        if (cached) return cached;
        throw err;
      }
    },
    async upsert(entry: LocationEntry): Promise<void> {
      updateCacheSingle('location_entries', entry);
      try {
        const { error } = await supabase.from('location_entries').upsert(entry);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: location entry upsert saved locally');
      }
    },
    async upsertAll(entries: LocationEntry[]): Promise<void> {
      if (!entries.length) return;
      const payload1 = entries.map(e => ({
        id: e.id,
        qrcode: e.qrcode,
        sku: e.sku,
        partner: e.partner,
        date: e.date,
        location: e.location,
        note: e.note,
        quantity: e.quantity,
        type: e.type,
        scanType: e.scanType,
        created_at: e.created_at
      }));
      
      setCache('location_entries', entries, 30 * 24 * 3600000);
      try {
        const { error } = await supabase.from('location_entries').upsert(payload1);
        
        if (error) {
          // Fallback for scantype (lowercase in older PostgREST schemas)
          const payload2 = entries.map(e => ({
            id: e.id,
            qrcode: e.qrcode,
            sku: e.sku,
            partner: e.partner,
            date: e.date,
            location: e.location,
            note: e.note,
            quantity: e.quantity,
            type: e.type,
            scantype: e.scanType,
            created_at: e.created_at
          }));
          const retry = await supabase.from('location_entries').upsert(payload2);
          if (retry.error) {
            throw retry.error;
          }
        }
      } catch (err) {
        console.warn('Offline mode: upsertAll location entries saved locally');
      }
    },
    async delete(id: string): Promise<void> {
      const currentEntries = getCache<LocationEntry[]>('location_entries') || [];
      setCache('location_entries', currentEntries.filter(e => e.id !== id), 30 * 24 * 3600000);
      
      try {
        const { error } = await supabase.from('location_entries').delete().eq('id', id);
        if (error) {
          console.error('Database error during location entry deletion:', error);
          throw error;
        }
      } catch (err) {
        console.warn('Offline mode: location entry deletion saved locally');
      }
    },
    async deleteMany(ids: string[]): Promise<void> {
      if (!ids.length) return;
      const currentEntries = getCache<LocationEntry[]>('location_entries') || [];
      setCache('location_entries', currentEntries.filter(e => !ids.includes(e.id)), 30 * 24 * 3600000);
      
      try {
        const chunks = chunkArray(ids, 100);
        for (const chunk of chunks) {
          const { error } = await supabase.from('location_entries').delete().in('id', chunk);
          if (error) {
            console.error('Database error during location entries bulk deletion:', error);
            throw error;
          }
          await delay(50);
        }
      } catch (err) {
        console.warn('Offline mode: location entries bulk deletion saved locally');
      }
    },
    async deleteByQRCode(qrcode: string): Promise<void> {
      const currentEntries = getCache<LocationEntry[]>('location_entries') || [];
      setCache('location_entries', currentEntries.filter(e => e.qrcode !== qrcode), 30 * 24 * 3600000);
      
      try {
        const { error } = await supabase.from('location_entries').delete().eq('qrcode', qrcode);
        if (error) {
          console.error('Database error during location entries deletion by qrcode:', error);
          throw error;
        }
      } catch (err) {
        console.warn('Offline mode: location entries deletion by qrcode saved locally');
      }
    },
    async deleteByQRCodes(qrcodes: string[]): Promise<void> {
      if (!qrcodes.length) return;
      const currentEntries = getCache<LocationEntry[]>('location_entries') || [];
      setCache('location_entries', currentEntries.filter(e => !qrcodes.includes(e.qrcode)), 30 * 24 * 3600000);
      
      try {
        const chunks = chunkArray(qrcodes, 100);
        for (const chunk of chunks) {
          const { error } = await supabase.from('location_entries').delete().in('qrcode', chunk);
          if (error) {
            console.error('Database error during location entries bulk deletion by qrcodes:', error);
            throw error;
          }
          await delay(50);
        }
      } catch (err) {
        console.warn('Offline mode: location entries bulk deletion by qrcodes saved locally');
      }
    },
    async deleteAll(): Promise<void> {
      setCache('location_entries', [], 30 * 24 * 3600000);
      const { error } = await supabase.from('location_entries').delete().neq('id', '');
      if (error) {
        console.error('Database error during location_entries bulk deletion:', error);
        throw error;
      }
    }
  },
  savedDeliveryNotes: {
    async getAll(): Promise<{id: string, date: string, items: DeliveryNoteItem[]}[]> {
      try {
        const { data, error } = await supabase.from('saved_delivery_notes').select('*');
        if (error) throw error;
        setCache('saved_delivery_notes', data as any[], 30 * 24 * 3600000);
        return data as {id: string, date: string, items: DeliveryNoteItem[]}[];
      } catch (err: any) {
        console.warn('Offline mode: using cached saved delivery notes', err.message);
        const cached = getCache<{id: string, date: string, items: DeliveryNoteItem[]}[]>('saved_delivery_notes');
        if (cached) return cached;
        throw err;
      }
    },
    async upsert(note: {id: string, date: string, items: DeliveryNoteItem[]}): Promise<void> {
      updateCacheSingle('saved_delivery_notes', note);
      try {
        const { error } = await supabase.from('saved_delivery_notes').upsert(note);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: saved delivery note upsert saved locally');
      }
    },
    async upsertAll(notes: {id: string, date: string, items: DeliveryNoteItem[]}[]): Promise<void> {
      setCache('saved_delivery_notes', notes, 30 * 24 * 3600000);
      try {
        const { error } = await supabase.from('saved_delivery_notes').upsert(notes);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: saved delivery notes upsertAll saved locally');
      }
    },
    async delete(id: string): Promise<void> {
      const currentNotes = getCache<{id: string, date: string, items: DeliveryNoteItem[]}[]>('saved_delivery_notes') || [];
      setCache('saved_delivery_notes', currentNotes.filter(n => n.id !== id), 30 * 24 * 3600000);
      
      const { error } = await supabase.from('saved_delivery_notes').delete().eq('id', id);
      if (error) {
        console.error('Database error during saved delivery note deletion:', error);
        throw error;
      }
    },
    async deleteMany(ids: string[]): Promise<void> {
      if (!ids.length) return;
      const currentNotes = getCache<{id: string, date: string, items: DeliveryNoteItem[]}[]>('saved_delivery_notes') || [];
      setCache('saved_delivery_notes', currentNotes.filter(n => !ids.includes(n.id)), 30 * 24 * 3600000);
      
      const chunks = chunkArray(ids, 100);
      for (const chunk of chunks) {
        const { error } = await supabase.from('saved_delivery_notes').delete().in('id', chunk);
        if (error) {
          console.error('Database error during saved delivery notes bulk deletion:', error);
          throw error;
        }
        await delay(50);
      }
    }
  },
  deliveryNoteHeader: {
    async get(): Promise<{docCode: string, dept: string, to: string, date: string} | null> {
      try {
        const { data, error } = await supabase.from('delivery_note_header').select('*').single();
        if (error && error.code !== 'PGRST116') throw error; // PGRST116 is "no rows returned"
        if (!data) return null;
        const result = { 
          docCode: data.doccode ?? data.docCode ?? '',
          dept: data.dept ?? '',
          to: data.toname ?? data.to ?? '',
          date: data.date ?? ''
        };
        setCache('delivery_note_header', result, 30 * 24 * 3600000);
        return result;
      } catch (err: any) {
        console.warn('Offline mode: using cached delivery note header', err.message);
        const cached = getCache<{docCode: string, dept: string, to: string, date: string}>('delivery_note_header');
        if (cached) return cached;
        throw err;
      }
    },
    async upsert(header: {docCode: string, dept: string, to: string, date: string}): Promise<void> {
      try {
        const { error } = await supabase.from('delivery_note_header').upsert({ 
          id: 'current', 
          doccode: header.docCode,
          dept: header.dept,
          toname: header.to,
          date: header.date
        });
        if (error) throw error;
        setCache('delivery_note_header', header, 30 * 24 * 3600000);
      } catch (error: any) {
        // Fallback in case they already ran the migration (or the error code differs)
        setCache('delivery_note_header', header, 30 * 24 * 3600000);
        try {
          const retry = await supabase.from('delivery_note_header').upsert({ 
            id: 'current', 
            doccode: header.docCode,
            dept: header.dept,
            to: header.to,
            date: header.date
          });
          if (retry.error) {
            throw retry.error;
          }
        } catch (err) {
          console.warn('Offline mode: delivery note header upsert saved locally');
        }
      }
    }
  },
  async getFullBackupData() {
    try {
      const [
        products,
        transactions,
        customers,
        deliveryNotes,
        locationEntries,
        savedDeliveryNotes,
        header
      ] = await Promise.all([
        this.products.getAll(),
        this.transactions.getAll(),
        this.customers.getAll(),
        this.deliveryNotes.getAll(),
        this.locationEntries.getAll(),
        this.savedDeliveryNotes.getAll(),
        this.deliveryNoteHeader.get()
      ]);

      return {
        products,
        transactions,
        customers,
        deliveryNotes,
        locationEntries,
        savedDeliveryNotes,
        header,
        backupDate: new Date().toISOString(),
        version: '1.0.0'
      };
    } catch (error) {
      console.error('Error fetching full backup data:', error);
      throw error;
    }
  },
  getBackupData() {
    const backup: Record<string, any> = {};
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.startsWith(CACHE_KEY_PREFIX)) {
        const data = localStorage.getItem(key);
        if (data) {
          try {
            backup[key] = JSON.parse(data);
          } catch (e) {
            backup[key] = data;
          }
        }
      }
    }
    return backup;
  },
  clearCache() {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.startsWith(CACHE_KEY_PREFIX)) {
        localStorage.removeItem(key);
        i--; // Adjust index after removal
      }
    }
  }
};
