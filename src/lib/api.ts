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
    async delete(id: string): Promise<void> {
      const { error } = await supabase.from('location_entries').delete().eq('id', id);
      if (error) throw error;
      localStorage.removeItem(CACHE_KEY_PREFIX + 'location_entries');
    }
  }
};
