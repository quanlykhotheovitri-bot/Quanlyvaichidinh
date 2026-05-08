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

const autoFreeStorage = () => {
  try {
    let total = 0;
    const items: { key: string; size: number }[] = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.startsWith(CACHE_KEY_PREFIX)) {
        const val = localStorage.getItem(key) || '';
        const size = (key.length + val.length) * 2;
        total += size;
        items.push({ key, size });
      }
    }

    // If over 4MB, remove the largest items until under 2MB
    if (total > 4 * 1024 * 1024) {
      items.sort((a, b) => b.size - a.size);
      while (total > 2 * 1024 * 1024 && items.length > 0) {
        const item = items.shift();
        if (item) {
          localStorage.removeItem(item.key);
          total -= item.size;
        }
      }
    }
  } catch (e) {
    console.error('Error in autoFreeStorage:', e);
  }
};

const setCache = <T>(key: string, value: T, ttl: number = 3600000): void => {
  try {
    autoFreeStorage();
    const expiry = Date.now() + ttl;
    localStorage.setItem(CACHE_KEY_PREFIX + key, JSON.stringify({ value, expiry }));
  } catch (e) {
    console.warn('Storage limit reached, attempting to free oldest cache items');
    // If setting fails, remove the oldest item and try again
    try {
      const items: { key: string; expiry: number }[] = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && k.startsWith(CACHE_KEY_PREFIX)) {
          const val = localStorage.getItem(k) || '';
          const { expiry } = JSON.parse(val);
          items.push({ key: k, expiry });
        }
      }
      items.sort((a, b) => a.expiry - b.expiry);
      if (items.length > 0) {
        localStorage.removeItem(items[0].key);
        const expiry = Date.now() + ttl;
        localStorage.setItem(CACHE_KEY_PREFIX + key, JSON.stringify({ value, expiry }));
      }
    } catch (err) {
      console.error('Final cache failure - local storage possibly full:', err);
    }
  }
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

const updateCacheMultiple = <T extends {id: string}>(key: string, items: T[]) => {
  const current = getCache<T[]>(key) || [];
  const currentMap = new Map(current.map(item => [item.id, item]));
  items.forEach(item => {
    currentMap.set(item.id, item);
  });
  setCache(key, Array.from(currentMap.values()), 30 * 24 * 3600000);
};

const chunkArray = <T>(array: T[], size: number): T[][] => {
  const result: T[][] = [];
  for (let i = 0; i < array.length; i += size) {
    result.push(array.slice(i, i + size));
  }
  return result;
};

const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

const fetchAll = async <T>(table: string): Promise<T[]> => {
  let allData: T[] = [];
  let from = 0;
  const step = 1000;
  let hasMore = true;

  while (hasMore) {
    const { data, error } = await supabase
      .from(table)
      .select('*')
      .range(from, from + step - 1);

    if (error) throw error;
    if (data && data.length > 0) {
      allData = [...allData, ...data];
      if (data.length < step) {
        hasMore = false;
      } else {
        from += step;
      }
    } else {
      hasMore = false;
    }
  }
  return allData;
};

export const api = {
  products: {
    async getAll(): Promise<Product[]> {
      try {
        const data = await fetchAll<any>('products');
        
        const mappedData = (data || []).map(row => ({
          ...row,
          minStock: Number(row.minstock ?? row.minStock ?? 0),
          lotNo: row.lotno ?? row.lotNo ?? '',
          ghiChu: row.ghichu ?? row.ghiChu ?? '',
          designationCode: row.designationcode ?? row.designationCode ?? '',
          loaiChiDinh: row.loaichidinh ?? row.loaiChiDinh ?? ''
        })) as Product[];

        // Smart Cache Merge
        const cached = getCache<Product[]>('products');
        if ((!mappedData || mappedData.length === 0) && cached && cached.length > 0) {
          console.warn('Server returned empty products, preserving local cache to prevent data loss');
          return cached;
        }

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
        id: product.id,
        sku: product.sku,
        name: product.name,
        category: product.category,
        unit: product.unit,
        minstock: product.minStock || 0,
        lotno: product.lotNo || '',
        ghichu: product.ghiChu || '',
        designationcode: product.designationCode || '',
        loaichidinh: product.loaiChiDinh || '',
        rpro: product.rpro || ''
      };

      try {
        const { error } = await supabase.from('products').upsert(payload);
        if (error) throw error;
      } catch (err) {
        console.warn('Product upsert failed, stored in local cache:', err);
      }
    },
    async upsertAll(products: Product[]): Promise<void> {
      if (!products.length) return;
      const payload = products.map(product => ({
        id: product.id,
        sku: product.sku,
        name: product.name,
        category: product.category,
        unit: product.unit,
        minstock: product.minStock || 0,
        lotno: product.lotNo || '',
        ghichu: product.ghiChu || '',
        designationcode: product.designationCode || '',
        loaichidinh: product.loaiChiDinh || '',
        rpro: product.rpro || ''
      }));
      updateCacheMultiple('products', products);
      try {
        const chunks = chunkArray(payload, 500);
        for (const chunk of chunks) {
          const { error } = await supabase.from('products').upsert(chunk);
          if (error) throw error;
          await delay(50);
        }
      } catch (err) {
        console.error('Error in products upsertAll:', err);
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
        const data = await fetchAll<any>('transactions');
        
        const mappedData = (data || []).map(row => ({
          ...row,
          productId: row.productid ?? row.productId ?? '',
          quantity: Number(row.quantity ?? row.qty ?? 0),
          loaiChiDinh: row.loaichidinh ?? row.loaiChiDinh ?? '',
          lotNo: row.lotno ?? row.lotNo ?? '',
          ghiChu: row.ghichu ?? row.ghiChu ?? '',
          designationCode: row.designationcode ?? row.designationCode ?? '',
          updateDate: row.updatedate ?? row.updateDate ?? '',
          isDeleted: Boolean(row.isdeleted ?? row.isDeleted ?? false)
        })) as Transaction[];

        // Smart Cache Merge: CRITICAL for preventing loss of imported data not yet on DB
        const cached = getCache<Transaction[]>('transactions');
        if ((!mappedData || mappedData.length === 0) && cached && cached.length > 0) {
          console.warn('Server returned empty transactions, preserving local cache to prevent data loss');
          return cached;
        }

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
        id: transaction.id,
        productid: transaction.productId,
        type: transaction.type,
        quantity: Number(transaction.quantity) || 0,
        date: transaction.date,
        partner: transaction.partner || 'N/A',
        lotno: transaction.lotNo || '',
        loaichidinh: transaction.loaiChiDinh || '',
        ghichu: transaction.ghiChu || '',
        designationcode: transaction.designationCode || '',
        rpro: transaction.rpro || '',
        updatedate: transaction.updateDate || '',
        isdeleted: transaction.isDeleted || false
      };

      try {
        const { error } = await supabase.from('transactions').upsert(payload);
        if (error) throw error;
      } catch (err) {
        console.warn('Transaction upsert failed, stored in local cache:', err);
      }
    },
    async upsertAll(transactions: Transaction[]): Promise<void> {
      if (!transactions.length) return;
      const payload = transactions.map(transaction => ({
        id: transaction.id,
        productid: transaction.productId,
        type: transaction.type,
        quantity: Number(transaction.quantity) || 0,
        date: transaction.date,
        partner: transaction.partner || 'N/A',
        lotno: transaction.lotNo || '',
        loaichidinh: transaction.loaiChiDinh || '',
        ghichu: transaction.ghiChu || '',
        designationcode: transaction.designationCode || '',
        rpro: transaction.rpro || '',
        updatedate: transaction.updateDate || '',
        isdeleted: transaction.isDeleted || false
      }));
      
      updateCacheMultiple('transactions', transactions);
      
      try {
        const chunks = chunkArray(payload, 500);
        for (const chunk of chunks) {
          const { error } = await supabase.from('transactions').upsert(chunk);
          if (error) throw error;
          await delay(50);
        }
      } catch (err) {
        console.error('Error in transactions upsertAll:', err);
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
        const data = await fetchAll<any>('customers');
        
        const mappedData = data as Customer[];
        
        // Smart Cache Merge
        const cached = getCache<Customer[]>('customers');
        if ((!mappedData || mappedData.length === 0) && cached && cached.length > 0) {
          console.warn('Server returned empty customers, preserving local cache');
          return cached;
        }

        setCache('customers', mappedData, 30 * 24 * 3600000);
        return mappedData;
      } catch (err: any) {
        console.warn('Offline mode: using cached customers', err.message);
        const cached = getCache<Customer[]>('customers');
        if (cached) return cached;
        throw err;
      }
    },
    async upsert(customer: Customer): Promise<void> {
      updateCacheSingle('customers', customer);
      const payload = {
        id: customer.id,
        code: customer.code,
        name: customer.name
      };
      try {
        const { error } = await supabase.from('customers').upsert(payload);
        if (error) throw error;
      } catch (err) {
        console.warn('Customer upsert saved locally:', err);
      }
    },
    async upsertAll(customers: Customer[]): Promise<void> {
      if (!customers.length) return;
      const payload = customers.map(c => ({
        id: c.id,
        code: c.code,
        name: c.name
      }));
      updateCacheMultiple('customers', customers);
      try {
        const chunks = chunkArray(payload, 500);
        for (const chunk of chunks) {
          const { error } = await supabase.from('customers').upsert(chunk);
          if (error) throw error;
          await delay(50);
        }
      } catch (err) {
        console.error('Error in customers upsertAll:', err);
        console.warn('Sync failed: data remains in local cache');
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
    },
    async deleteAll(): Promise<void> {
      setCache('customers', [], 30 * 24 * 3600000);
      const { error } = await supabase.from('customers').delete().neq('id', '');
      if (error) {
        console.error('Database error during customers bulk deletion:', error);
        throw error;
      }
    }
  },
  deliveryNotes: {
    async getAll(): Promise<DeliveryNoteItem[]> {
      try {
        const data = await fetchAll<any>('delivery_notes');
        
        const mappedData = (data || []).map(row => ({
          ...row,
          ovnSaleOrder: row.ovnsaleorder ?? row.ovnSaleOrder ?? '',
          ovnProductionOrder: row.ovnproductionorder ?? row.ovnProductionOrder ?? '',
          materialName: row.materialname ?? row.materialName ?? '',
          qtyErp: Number(row.qtyerp ?? row.qtyErp ?? 0),
          actualQty: Number(row.actualqty ?? row.actualQty ?? 0),
          lotNo: row.lotno ?? row.lotNo ?? '',
          actualIssuedQty: Number(row.actualissuedqty ?? row.actualIssuedQty ?? 0),
          customerCode: row.customercode ?? row.customerCode ?? '',
          finalDestination: row.finaldestination ?? row.finalDestination ?? '',
          noCode: row.nocode ?? row.noCode ?? '',
          loaiChiDinh: row.loaichidinh ?? row.loaiChiDinh ?? '',
          assignedLots: row.assignedlots ?? row.assignedLots ?? []
        })) as DeliveryNoteItem[];

        // Smart Cache Merge
        const cached = getCache<DeliveryNoteItem[]>('delivery_notes');
        if ((!mappedData || mappedData.length === 0) && cached && cached.length > 0) {
          console.warn('Server returned empty delivery notes, preserving local cache');
          return cached;
        }

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
      const payload = items.map(item => ({
        id: item.id,
        no: item.no,
        ovnsaleorder: item.ovnSaleOrder || '',
        ovnproductionorder: item.ovnProductionOrder || '',
        item: item.item || '',
        materialname: item.materialName || '',
        unit: item.unit || '',
        qtyerp: item.qtyErp || 0,
        actualqty: item.actualQty || 0,
        lotno: item.lotNo || '',
        actualissuedqty: item.actualIssuedQty || 0,
        remark: item.remark || '',
        brand: item.brand || '',
        customercode: item.customerCode || '',
        finaldestination: item.finalDestination || '',
        nocode: item.noCode || '',
        location: item.location || '',
        stock: item.stock || '',
        loaichidinh: item.loaiChiDinh || '',
        assignedlots: item.assignedLots || []
      }));
      updateCacheMultiple('delivery_notes', items);
      try {
        const chunks = chunkArray(payload, 500);
        for (const chunk of chunks) {
          const { error } = await supabase.from('delivery_notes').upsert(chunk);
          if (error) throw error;
          await delay(50);
        }
      } catch (err) {
        console.error('Error in deliveryNotes upsertAll:', err);
        console.warn('Sync failed: data remains in local cache');
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
        const data = await fetchAll<any>('location_entries');
        
        const mappedData = (data || []).map(row => ({
          ...row,
          quantity: Number(row.quantity ?? 1),
          scanType: row.scantype ?? row.scanType ?? 'INPUT'
        })) as LocationEntry[];

        // Smart Cache Merge
        const cached = getCache<LocationEntry[]>('location_entries');
        if ((!mappedData || mappedData.length === 0) && cached && cached.length > 0) {
          console.warn('Server returned empty location entries, preserving local cache');
          return cached;
        }

        setCache('location_entries', mappedData, 30 * 24 * 3600000);
        return mappedData;
      } catch (err: any) {
        console.warn('Offline mode: using cached location entries', err.message);
        const cached = getCache<LocationEntry[]>('location_entries');
        if (cached) return cached;
        throw err;
      }
    },
    async upsert(entry: LocationEntry): Promise<void> {
      updateCacheSingle('location_entries', entry);
      const payload = {
        id: entry.id,
        qrcode: entry.qrcode,
        sku: entry.sku,
        partner: entry.partner || '',
        date: entry.date || '',
        location: entry.location || '',
        note: entry.note || '',
        quantity: entry.quantity || 1,
        type: entry.type,
        scantype: entry.scanType || 'INPUT'
      };
      try {
        const { error } = await supabase.from('location_entries').upsert(payload);
        if (error) throw error;
      } catch (err) {
        console.warn('Location entry sync failed, saved locally:', err);
      }
    },
    async upsertAll(entries: LocationEntry[]): Promise<void> {
      if (!entries.length) return;
      const payload = entries.map(e => ({
        id: e.id,
        qrcode: e.qrcode,
        sku: e.sku,
        partner: e.partner || '',
        date: e.date || '',
        location: e.location || '',
        note: e.note || '',
        quantity: e.quantity || 1,
        type: e.type,
        scantype: e.scanType || 'INPUT'
      }));
      
      updateCacheMultiple('location_entries', entries);
      try {
        const chunks = chunkArray(payload, 500);
        for (const chunk of chunks) {
          const { error } = await supabase.from('location_entries').upsert(chunk);
          if (error) throw error;
          await delay(50);
        }
      } catch (err) {
        console.error('Error in locationEntries upsertAll:', err);
        console.warn('Sync failed: data remains in local cache');
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
    async getAll(): Promise<{id: string, date: string, dept: string, items: DeliveryNoteItem[]}[]> {
      try {
        const data = await fetchAll<any>('saved_delivery_notes');
        
        const mappedData = (data || []).map(note => ({
          ...note,
          dept: note.dept || 'SX 5'
        })) as {id: string, date: string, dept: string, items: DeliveryNoteItem[]}[];

        // Smart Cache Merge
        const cached = getCache<{id: string, date: string, dept: string, items: DeliveryNoteItem[]}[]>('saved_delivery_notes');
        if ((!mappedData || mappedData.length === 0) && cached && cached.length > 0) {
          console.warn('Server returned empty saved delivery notes, preserving local cache');
          return cached;
        }

        setCache('saved_delivery_notes', mappedData, 30 * 24 * 3600000);
        return mappedData;
      } catch (err: any) {
        console.warn('Offline mode: using cached saved delivery notes', err.message);
        const cached = getCache<{id: string, date: string, dept: string, items: DeliveryNoteItem[]}[]>('saved_delivery_notes');
        if (cached) return cached;
        throw err;
      }
    },
    async upsert(note: {id: string, date: string, dept: string, items: DeliveryNoteItem[]}): Promise<void> {
      updateCacheSingle('saved_delivery_notes', note);
      try {
        const { error } = await supabase.from('saved_delivery_notes').upsert(note);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: saved delivery note upsert saved locally');
      }
    },
    async upsertAll(notes: {id: string, date: string, dept: string, items: DeliveryNoteItem[]}[]): Promise<void> {
      setCache('saved_delivery_notes', notes, 30 * 24 * 3600000);
      try {
        const { error } = await supabase.from('saved_delivery_notes').upsert(notes);
        if (error) throw error;
      } catch (err) {
        console.warn('Offline mode: saved delivery notes upsertAll saved locally');
      }
    },
    async delete(id: string): Promise<void> {
      const currentNotes = getCache<{id: string, date: string, dept: string, items: DeliveryNoteItem[]}[]>('saved_delivery_notes') || [];
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
    },
    async deleteAll(): Promise<void> {
      setCache('saved_delivery_notes', [], 30 * 24 * 3600000);
      const { error } = await supabase.from('saved_delivery_notes').delete().neq('id', '');
      if (error) {
        console.error('Database error during saved delivery notes bulk deletion:', error);
        throw error;
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
  async getDatabaseUsage() {
    try {
      const tables = ['products', 'transactions', 'customers', 'delivery_notes', 'location_entries', 'saved_delivery_notes'];
      let totalBytes = 0;
      
      // We use the cached data to estimate size to avoid heavy network calls
      // This is a good proxy since the app caches all main tables
      for (const table of tables) {
        const cached = localStorage.getItem(CACHE_KEY_PREFIX + table);
        if (cached) {
          totalBytes += cached.length * 2; // UTF-16 characters are 2 bytes
        }
      }
      
      return totalBytes;
    } catch (error) {
      console.error('Error estimating database usage:', error);
      return 0;
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
