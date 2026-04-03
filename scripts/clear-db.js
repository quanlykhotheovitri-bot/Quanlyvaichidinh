import { createClient } from '@supabase/supabase-js';
import * as dotenv from 'dotenv';

// Load environment variables
dotenv.config();

const supabaseUrl = process.env.VITE_SUPABASE_URL;
const supabaseKey = process.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseKey) {
  console.error('Error: VITE_SUPABASE_URL or VITE_SUPABASE_ANON_KEY not found in environment.');
  console.log('Please make sure your .env file is set up correctly.');
  process.exit(1);
}

const supabase = createClient(supabaseUrl, supabaseKey);

async function clearTable(tableName) {
  console.log(`Cleaning table: ${tableName}...`);
  const { error } = await supabase.from(tableName).delete().neq('id', '');
  if (error) {
    console.error(`Error cleaning ${tableName}:`, error.message);
  } else {
    console.log(`Successfully cleaned ${tableName}.`);
  }
}

async function startCleanup() {
  const tables = [
    'products',
    'transactions',
    'customers',
    'location_entries',
    'delivery_notes',
    'saved_delivery_notes',
    'delivery_note_header'
  ];

  console.log('--- STARTING SUPABASE DATABASE CLEANUP ---');
  for (const table of tables) {
    await clearTable(table);
  }
  console.log('--- CLEANUP COMPLETE ---');
  console.log('You can now refresh your app (F5) to see the empty state.');
}

startCleanup();
