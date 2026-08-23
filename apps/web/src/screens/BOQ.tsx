/**
 * BOQTab — BOQ = scope structure (NOT commercial authority).
 *
 * After every mutation, re-fetch authoritative state. The browser never
 * computes commercial totals. (Phase 2C.1 §10, §14)
 */
import { useEffect, useState } from 'react'
import { boqApi, type BOQ, type BOQItem } from '../api/client'
import { styles } from '../styles'

export function BOQTab({ projectId }: { projectId: string }) {
  const [boqs, setBoqs] = useState<BOQ[]>([])
  const [selected, setSelected] = useState<BOQ | null>(null)
  const [items, setItems] = useState<BOQItem[]>([])
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)
  const [newName, setNewName] = useState('')
  // new item form
  const [itemCode, setItemCode] = useState('')
  const [itemDesc, setItemDesc] = useState('')
  const [itemUnit, setItemUnit] = useState('m2')
  const [itemQty, setItemQty] = useState(0)

  const refreshBoqs = async () => {
    try {
      setBoqs(await boqApi.listForProject(projectId))
    } catch (e) {
      setError(String(e))
    }
  }
  const refreshItems = async (boqId: string) => {
    try {
      setItems(await boqApi.listItems(boqId))
    } catch (e) {
      setError(String(e))
    }
  }

  useEffect(() => {
    refreshBoqs()
  }, [projectId])

  const selectBoq = async (b: BOQ) => {
    setSelected(b)
    setLoading(true)
    setError(null)
    await refreshItems(b.boqId)
    setLoading(false)
  }

  const createBoq = async () => {
    setError(null)
    try {
      await boqApi.create(projectId, newName || undefined)
      setNewName('')
      await refreshBoqs()
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Create failed')
    }
  }

  const addItem = async () => {
    if (!selected || !itemCode || !itemDesc) return
    setError(null)
    try {
      await boqApi.addItem(selected.boqId, {
        itemCode,
        description: itemDesc,
        unit: itemUnit,
        quantityValue: itemQty,
        quantityUnit: itemUnit,
        provenance: 'manual',
      })
      setItemCode('')
      setItemDesc('')
      setItemQty(0)
      await refreshItems(selected.boqId)
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Add failed')
    }
  }

  const updateQty = async (itemId: string, qty: number) => {
    if (!selected) return
    setError(null)
    try {
      await boqApi.updateQuantity(itemId, qty, itemUnit)
      await refreshItems(selected.boqId)
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Update failed')
    }
  }

  return (
    <div style={styles.column}>
      {error && <div style={styles.error}>{error}</div>}
      <div style={styles.card}>
        <h2 style={styles.title}>BOQs</h2>
        <p style={styles.muted}>BOQ = scope structure, not commercial authority.</p>
        <div style={styles.row}>
          <input
            style={styles.input}
            placeholder="BOQ name (optional)"
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
          />
          <button style={styles.buttonPrimary} onClick={createBoq}>
            Create BOQ
          </button>
        </div>
        <div style={styles.column}>
          {boqs.map((b) => (
            <button
              key={b.boqId}
              style={selected?.boqId === b.boqId ? styles.buttonPrimary : styles.button}
              onClick={() => selectBoq(b)}
            >
              {b.boqId.slice(-8)}
            </button>
          ))}
        </div>
      </div>
      {selected && (
        <div style={styles.card}>
          <h3 style={styles.title}>Items in {selected.boqId.slice(-8)}</h3>
          {loading ? (
            <div style={styles.loading}>Loading…</div>
          ) : items.length === 0 ? (
            <p style={styles.muted}>No items yet.</p>
          ) : (
            <table style={styles.table}>
              <thead>
                <tr>
                  <th style={styles.th}>Code</th>
                  <th style={styles.th}>Description</th>
                  <th style={styles.th}>Unit</th>
                  <th style={styles.th}>Quantity</th>
                  <th style={styles.th}></th>
                </tr>
              </thead>
              <tbody>
                {items.map((it) => (
                  <tr key={it.itemId}>
                    <td style={styles.td}>{it.itemCode}</td>
                    <td style={styles.td}>{it.description}</td>
                    <td style={styles.td}>{it.quantity.unit}</td>
                    <td style={styles.td}>{it.quantity.value}</td>
                    <td style={styles.td}>
                      <input
                        type="number"
                        style={{ ...styles.input, width: 80 }}
                        defaultValue={it.quantity.value}
                        onBlur={(e) => {
                          const v = Number(e.target.value)
                          if (v !== it.quantity.value) updateQty(it.itemId, v)
                        }}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
          <h3 style={styles.title}>Add item</h3>
          <div style={styles.row}>
            <input
              style={styles.input}
              placeholder="Code"
              value={itemCode}
              onChange={(e) => setItemCode(e.target.value)}
            />
            <input
              style={styles.input}
              placeholder="Description"
              value={itemDesc}
              onChange={(e) => setItemDesc(e.target.value)}
            />
            <input
              style={{ ...styles.input, width: 80 }}
              placeholder="Unit"
              value={itemUnit}
              onChange={(e) => setItemUnit(e.target.value)}
            />
            <input
              type="number"
              style={{ ...styles.input, width: 80 }}
              placeholder="Qty"
              value={itemQty}
              onChange={(e) => setItemQty(Number(e.target.value))}
            />
            <button
              style={styles.buttonPrimary}
              onClick={addItem}
              disabled={!itemCode || !itemDesc}
            >
              Add item
            </button>
          </div>
        </div>
      )}
    </div>
  )
}
