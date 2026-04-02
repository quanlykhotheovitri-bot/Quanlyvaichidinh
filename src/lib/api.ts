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
      return data as Product[];
    },
    async upsert(product: Product): Promise<void> {
      const { error } = await supabase.from('products').upsert(product);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'products');
    },
    async upsertAll(products: Product[]): Promise<void> {
      const { error } = await supabase.from('products').upsert(products);
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
      return data as Transaction[];
    },
    async upsert(transaction: Transaction): Promise<void> {
      const { error } = await supabase.from('transactions').upsert(transaction);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'transactions');
    },
    async upsertAll(transactions: Transaction[]): Promise<void> {
      const { error } = await supabase.from('transactions').upsert(transactions);
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
      return data as DeliveryNoteItem[];
    },
    async upsertAll(items: DeliveryNoteItem[]): Promise<void> {
      if (items.length === 0) return;
      const { error } = await supabase.from('delivery_notes').upsert(items);
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
      if (entries.length === 0) return;
      const { error } = await supabase.from('location_entries').upsert(entries);
      if (error) throw error;
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
      return { ...data, to: data.to || (data as any).toName || '' };
    },
    async upsert(header: {docCode: string, dept: string, to: string, date: string}): Promise<void> {
      const { error } = await supabase.from('delivery_note_header').upsert({ 
        id: 'current', 
        docCode: header.docCode,
        dept: header.dept,
        to: header.to,
        date: header.date
      });
      if (error) {
        // Fallback in case they already ran the migration (or the error code differs)
        const retry = await supabase.from('delivery_note_header').upsert({ 
          id: 'current', 
          docCode: header.docCode,
          dept: header.dept,
          toName: header.to,
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
