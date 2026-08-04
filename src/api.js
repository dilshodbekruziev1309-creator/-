import { supabase } from './supabaseClient'

async function handle(promise) {
  const { data, error } = await promise
  if (error) throw error
  return data
}

export const api = {
  products: {
    list: () => handle(supabase.from('products').select('*').order('created_at')),
    add: (p) => handle(supabase.from('products').insert(p).select().single()),
    update: (id, patch) => handle(supabase.from('products').update(patch).eq('id', id).select().single()),
    remove: (id) => handle(supabase.from('products').delete().eq('id', id)),
  },
  customers: {
    list: () => handle(supabase.from('customers').select('*').order('created_at')),
    add: (c) => handle(supabase.from('customers').insert(c).select().single()),
    update: (id, patch) => handle(supabase.from('customers').update(patch).eq('id', id).select().single()),
    remove: (id) => handle(supabase.from('customers').delete().eq('id', id)),
  },
  invoices: {
    list: () => handle(supabase.from('invoices').select('*').order('created_at')),
    add: (i) => handle(supabase.from('invoices').insert(i).select().single()),
    update: (id, patch) => handle(supabase.from('invoices').update(patch).eq('id', id).select().single()),
    remove: (id) => handle(supabase.from('invoices').delete().eq('id', id)),
  },
  payments: {
    list: () => handle(supabase.from('payments').select('*').order('created_at')),
    add: (p) => handle(supabase.from('payments').insert(p).select().single()),
    remove: (id) => handle(supabase.from('payments').delete().eq('id', id)),
  },
  stockMovements: {
    list: () => handle(supabase.from('stock_movements').select('*').order('date')),
    add: (m) => handle(supabase.from('stock_movements').insert(m).select().single()),
  },
  finishedProducts: {
    list: () => handle(supabase.from('finished_products').select('*').order('created_at')),
    add: (f) => handle(supabase.from('finished_products').insert(f).select().single()),
    update: (id, patch) => handle(supabase.from('finished_products').update(patch).eq('id', id).select().single()),
    remove: (id) => handle(supabase.from('finished_products').delete().eq('id', id)),
  },
  rawMaterials: {
    list: () => handle(supabase.from('raw_materials').select('*').order('created_at')),
    add: (m) => handle(supabase.from('raw_materials').insert(m).select().single()),
    remove: (id) => handle(supabase.from('raw_materials').delete().eq('id', id)),
  },
  rawMaterialBatches: {
    list: () => handle(supabase.from('raw_material_batches').select('*').order('date')),
    add: (b) => handle(supabase.from('raw_material_batches').insert(b).select().single()),
    update: (id, patch) => handle(supabase.from('raw_material_batches').update(patch).eq('id', id).select().single()),
    remove: (id) => handle(supabase.from('raw_material_batches').delete().eq('id', id)),
  },
  productNorms: {
    list: () => handle(supabase.from('product_norms').select('*')),
    add: (n) => handle(supabase.from('product_norms').insert(n).select().single()),
    remove: (id) => handle(supabase.from('product_norms').delete().eq('id', id)),
  },
  productionBatches: {
    list: () => handle(supabase.from('production_batches').select('*').order('created_at')),
    add: (b) => handle(supabase.from('production_batches').insert(b).select().single()),
    update: (id, patch) => handle(supabase.from('production_batches').update(patch).eq('id', id).select().single()),
    remove: (id) => handle(supabase.from('production_batches').delete().eq('id', id)),
  },
  productionConsumptions: {
    list: () => handle(supabase.from('production_consumptions').select('*')),
    add: (c) => handle(supabase.from('production_consumptions').insert(c).select().single()),
    removeByBatch: (productionBatchId) => handle(supabase.from('production_consumptions').delete().eq('production_batch_id', productionBatchId)),
  },
  cashTransactions: {
    list: () => handle(supabase.from('cash_transactions').select('*').order('date')),
    add: (t) => handle(supabase.from('cash_transactions').insert(t).select().single()),
    remove: (id) => handle(supabase.from('cash_transactions').delete().eq('id', id)),
  },
  cashCategories: {
    list: () => handle(supabase.from('cash_categories').select('*').order('name')),
    add: (c) => handle(supabase.from('cash_categories').insert(c).select().single()),
  },
  suppliers: {
    list: () => handle(supabase.from('suppliers').select('*').order('created_at')),
    add: (s) => handle(supabase.from('suppliers').insert(s).select().single()),
    update: (id, patch) => handle(supabase.from('suppliers').update(patch).eq('id', id).select().single()),
    remove: (id) => handle(supabase.from('suppliers').delete().eq('id', id)),
  },
  supplierPayments: {
    list: () => handle(supabase.from('supplier_payments').select('*').order('date')),
    add: (p) => handle(supabase.from('supplier_payments').insert(p).select().single()),
    remove: (id) => handle(supabase.from('supplier_payments').delete().eq('id', id)),
  },
  fixedAssets: {
    list: () => handle(supabase.from('fixed_assets').select('*').order('created_at')),
    add: (a) => handle(supabase.from('fixed_assets').insert(a).select().single()),
    update: (id, patch) => handle(supabase.from('fixed_assets').update(patch).eq('id', id).select().single()),
    remove: (id) => handle(supabase.from('fixed_assets').delete().eq('id', id)),
  },
  settings: {
    get: async () => {
      const { data } = await supabase.from('settings').select('*').eq('id', 1).maybeSingle()
      return data || { id: 1, company_name: 'Mening korxonam', company_phone: '' }
    },
    save: (s) => handle(supabase.from('settings').upsert({ id: 1, ...s }).select().single()),
  },
}
