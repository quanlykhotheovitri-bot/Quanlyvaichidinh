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

export const api = {
  products: {
    async getAll(): Promise<Product[]> {
      const cached = getCache<Product[]>('products');
      if (cached) return cached;

      const { data, error } = await supabase.from('products').select('*');
      if (error) throw error;
      
      setCache('products', data);
      return (data || []).map(row => ({
        ...row,
        minStock: row.minstock ?? row.minStock ?? 0,
        lotNo: row.lotno ?? row.lotNo ?? '',
        ghiChu: row.ghichu ?? row.ghiChu ?? '',
        designationCode: row.designationcode ?? row.designationCode ?? '',
        loaiChiDinh: row.loaichidinh ?? row.loaiChiDinh ?? ''
      })) as Product[];
    },
    async upsert(product: Product): Promise<void> {
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

      const { error } = await supabase.from('products').upsert(payload);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'products');
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
      const { error } = await supabase.from('products').upsert(payload);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'products');
    },
    async delete(id: string): Promise<void> {
      const { error } = await supabase.from('products').delete().eq('id', id);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'products');
    }
  },
  transactions: {
    async getAll(): Promise<Transaction[]> {
      const cached = getCache<Transaction[]>('transactions');
      if (cached) return cached;

      const { data, error } = await supabase.from('transactions').select('*');
      if (error) throw error;
      
      setCache('transactions', data);
      return (data || []).map(row => ({
        ...row,
        productId: row.productid ?? row.productId ?? '',
        loaiChiDinh: row.loaichidinh ?? row.loaiChiDinh ?? '',
        lotNo: row.lotno ?? row.lotNo ?? '',
        ghiChu: row.ghichu ?? row.ghiChu ?? '',
        designationCode: row.designationcode ?? row.designationCode ?? ''
      })) as Transaction[];
    },
    async upsert(transaction: Transaction): Promise<void> {
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

      const { error } = await supabase.from('transactions').upsert(payload);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'transactions');
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
      const { error } = await supabase.from('transactions').upsert(payload);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'transactions');
    },
    async delete(id: string): Promise<void> {
      const { error } = await supabase.from('transactions').delete().eq('id', id);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'transactions');
    }
  },
  customers: {
    async getAll(): Promise<Customer[]> {
      const cached = getCache<Customer[]>('customers');
      if (cached) return cached;

      const { data, error } = await supabase.from('customers').select('*');
      if (error) throw error;
      
      setCache('customers', data);
      return data as Customer[];
    },
    async upsert(customer: Customer): Promise<void> {
      const { error } = await supabase.from('customers').upsert(customer);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'customers');
    },
    async upsertAll(customers: Customer[]): Promise<void> {
      const { error } = await supabase.from('customers').upsert(customers);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'customers');
    },
    async delete(id: string): Promise<void> {
      const { error } = await supabase.from('customers').delete().eq('id', id);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'customers');
    }
  },
  deliveryNotes: {
    async getAll(): Promise<DeliveryNoteItem[]> {
      const cached = getCache<DeliveryNoteItem[]>('delivery_notes');
      if (cached) return cached;

      const { data, error } = await supabase.from('delivery_notes').select('*');
      if (error) throw error;
      
      setCache('delivery_notes', data);
      return (data || []).map(row => ({
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
      const { error } = await supabase.from('delivery_notes').upsert(payload);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'delivery_notes');
    },
    async deleteAll(): Promise<void> {
      const { error } = await supabase.from('delivery_notes').delete().neq('id', '');
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'delivery_notes');
    }
  },
  locationEntries: {
    async getAll(): Promise<LocationEntry[]> {
      const cached = getCache<LocationEntry[]>('location_entries');
      if (cached) return cached;

      const { data, error } = await supabase.from('location_entries').select('*');
      if (error) throw error;
      
      setCache('location_entries', data);
      return data as LocationEntry[];
    },
    async upsert(entry: LocationEntry): Promise<void> {
      const { error } = await supabase.from('location_entries').upsert(entry);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'location_entries');
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
        scanType: e.scanType
      }));
      
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
          scantype: e.scanType
        }));
        const retry = await supabase.from('location_entries').upsert(payload2);
        if (retry.error) {
          const errorMsg = `Lỗi Supabase (Location - Original): ${error.message} (Code: ${error.code})\nLỗi Supabase (Location - Retry): ${retry.error.message} (Code: ${retry.error.code})`;
          alert(errorMsg);
          throw retry.error;
        }
      }
      localStorage.removeItem(CACHE_KEY_PREFIX + 'location_entries');
    },
    async delete(id: string): Promise<void> {
      const { error } = await supabase.from('location_entries').delete().eq('id', id);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'location_entries');
    },
    async deleteByQRCode(qrcode: string): Promise<void> {
      const { error } = await supabase.from('location_entries').delete().eq('qrcode', qrcode);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'location_entries');
    }
  },
  savedDeliveryNotes: {
    async getAll(): Promise<{id: string, date: string, items: DeliveryNoteItem[]}[]> {
      const cached = getCache<{id: string, date: string, items: DeliveryNoteItem[]}[]>('saved_delivery_notes');
      if (cached) return cached;

      const { data, error } = await supabase.from('saved_delivery_notes').select('*');
      if (error) throw error;
      
      setCache('saved_delivery_notes', data);
      return data as {id: string, date: string, items: DeliveryNoteItem[]}[];
    },
    async upsert(note: {id: string, date: string, items: DeliveryNoteItem[]}): Promise<void> {
      const { error } = await supabase.from('saved_delivery_notes').upsert(note);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'saved_delivery_notes');
    },
    async upsertAll(notes: {id: string, date: string, items: DeliveryNoteItem[]}[]): Promise<void> {
      const { error } = await supabase.from('saved_delivery_notes').upsert(notes);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'saved_delivery_notes');
    },
    async delete(id: string): Promise<void> {
      const { error } = await supabase.from('saved_delivery_notes').delete().eq('id', id);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'saved_delivery_notes');
    }
  },
  deliveryNoteHeader: {
    async get(): Promise<{docCode: string, dept: string, to: string, date: string} | null> {
      const { data, error } = await supabase.from('delivery_note_header').select('*').single();
      if (error && error.code !== 'PGRST116') throw error; // PGRST116 is "no rows returned"
      if (!data) return null;
      return { 
        docCode: data.doccode ?? data.docCode ?? '',
        dept: data.dept ?? '',
        to: data.toname ?? data.to ?? '',
        date: data.date ?? ''
      };
    },
    async upsert(header: {docCode: string, dept: string, to: string, date: string}): Promise<void> {
      const { error } = await supabase.from('delivery_note_header').upsert({ 
        id: 'current', 
        doccode: header.docCode,
        dept: header.dept,
        toname: header.to,
        date: header.date
      });
      if (error) {
        // Fallback in case they already ran the migration (or the error code differs)
        const retry = await supabase.from('delivery_note_header').upsert({ 
          id: 'current', 
          doccode: header.docCode,
          dept: header.dept,
          to: header.to,
          date: header.date
        });
        if (retry.error) {
          // If both fail, let's show the exact error message from Supabase
          const errorMsg = `Lỗi Supabase (Original): ${error.message} (Code: ${error.code})\nLỗi Supabase (Retry): ${retry.error.message} (Code: ${retry.error.code})`;
          alert(errorMsg);
          throw retry.error;
        }
      }
    }
  }
};
