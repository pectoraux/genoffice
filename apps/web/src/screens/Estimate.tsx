/**
 * EstimateTab — EstimateRevision = canonical commercial authority.
 *
 * CRITICAL: the browser MUST NOT calculate overhead/contingency/totalCost/
 * profit/sellPrice/grossMargin. All derived totals come from the server's
 * replayEstimate response. The browser only edits payload INPUTS (line
 * descriptions, quantities, rates, policy) and sends them to the server.
 * (Phase 2C.1 §13, §14)
 *
 * After finalize: editing disabled, status=finalized, immutability is
 * server-enforced (the browser only reflects it). (§15)
 */
import { useEffect, useState, useCallback } from 'react'
import {
  estimateApi,
  buildEstimatePayload,
  type EstimateRevision,
  type EstimateReplay,
} from '../api/client'
import { styles } from '../styles'

export function EstimateTab({ projectId }: { projectId: string }) {
  const [revisions, setRevisions] = useState<EstimateRevision[]>([])
  const [selected, setSelected] = useState<EstimateRevision | null>(null)
  const [replay, setReplay] = useState<EstimateReplay | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [, setLoading] = useState(false)
  // draft input form
  const [currency, setCurrency] = useState('GHS')
  const [overheadPct, setOverheadPct] = useState(0.1)
  const [contingencyPct, setContingencyPct] = useState(0.05)
  const [profitMode, setProfitMode] = useState<'markup' | 'margin'>('markup')
  const [profitRatio, setProfitRatio] = useState(0.1)
  const [lineDesc, setLineDesc] = useState('Concrete')
  const [lineQty, setLineQty] = useState(100)
  const [lineUnit, setLineUnit] = useState('m2')
  const [lineRate, setLineRate] = useState(500)
  const [confirmingFinalize, setConfirmingFinalize] = useState(false)

  const refreshList = async () => {
    try {
      setRevisions(await estimateApi.listForProject(projectId))
    } catch (e) {
      setError(String(e))
    }
  }

  useEffect(() => {
    refreshList()
  }, [projectId])

  const selectRev = useCallback(async (rid: string) => {
    setLoading(true)
    setError(null)
    try {
      const r = await estimateApi.get(rid)
      setSelected(r)
      // populate the form from the loaded payload
      const p = r.payload
      setCurrency(p.currency)
      setOverheadPct(p.policy.overheadPct)
      setContingencyPct(p.policy.contingencyPct)
      setProfitMode(p.policy.targetProfitMode as 'markup' | 'margin')
      setProfitRatio(p.policy.targetProfitRatio)
      if (p.lines.length > 0) {
        const l0 = p.lines[0]!
        setLineDesc(l0.description)
        setLineQty(l0.quantity.value)
        setLineUnit(l0.quantity.unit)
        setLineRate(l0.rate?.amount ?? 0)
      }
      // fetch replay totals (authoritative — never computed client-side)
      try {
        setReplay(await estimateApi.replay(rid))
      } catch {
        setReplay(null)
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to load')
    } finally {
      setLoading(false)
    }
  }, [])

  const createDraft = async () => {
    setError(null)
    try {
      const payload = buildEstimatePayload({
        projectId,
        currency,
        overheadPct,
        contingencyPct,
        targetProfitMode: profitMode,
        targetProfitRatio: profitRatio,
        lines: [
          {
            lineId: 'l1',
            description: lineDesc,
            quantityValue: lineQty,
            quantityUnit: lineUnit,
            rateMinor: lineRate,
            costBasis: 'unit-rate',
            pricingStrategy: 'markup',
            pricingRatio: 0.2,
          },
        ],
        pricingAlgorithmVersion: 'v1',
      })
      const r = await estimateApi.create(projectId, payload)
      await refreshList()
      await selectRev(r.revisionId)
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Create failed')
    }
  }

  const saveDraft = async () => {
    if (!selected) return
    setError(null)
    try {
      const payload = buildEstimatePayload({
        projectId,
        currency,
        overheadPct,
        contingencyPct,
        targetProfitMode: profitMode,
        targetProfitRatio: profitRatio,
        lines: [
          {
            lineId: 'l1',
            description: lineDesc,
            quantityValue: lineQty,
            quantityUnit: lineUnit,
            rateMinor: lineRate,
            costBasis: 'unit-rate',
            pricingStrategy: 'markup',
            pricingRatio: 0.2,
          },
        ],
        pricingAlgorithmVersion: 'v1',
      })
      const r = await estimateApi.update(selected.revisionId, payload)
      setSelected(r)
      setReplay(await estimateApi.replay(r.revisionId))
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Update failed')
    }
  }

  const finalize = async () => {
    if (!selected) return
    setError(null)
    setConfirmingFinalize(false)
    try {
      const f = await estimateApi.finalize(selected.revisionId)
      setSelected(f)
      setReplay(await estimateApi.replay(f.revisionId))
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Finalize failed')
    }
  }

  const isFinalized = selected?.status === 'finalized' || selected?.status === 'superseded'
  const money = (m: { amount: number; currency: string } | null) =>
    m ? `${(m.amount / 100).toFixed(2)} ${m.currency}` : '—'

  return (
    <div style={styles.column}>
      {error && <div style={styles.error}>{error}</div>}
      <div style={styles.card}>
        <h2 style={styles.title}>Estimates</h2>
        <p style={styles.muted}>
          EstimateRevision = canonical commercial authority. Totals come from the server.
        </p>
        <div style={styles.row}>
          <button style={styles.buttonPrimary} onClick={createDraft}>
            Create draft
          </button>
          {revisions.map((r) => (
            <button
              key={r.revisionId}
              style={selected?.revisionId === r.revisionId ? styles.buttonPrimary : styles.button}
              onClick={() => selectRev(r.revisionId)}
            >
              rev {r.revisionNumber} ({r.status})
            </button>
          ))}
        </div>
      </div>

      {selected && (
        <div style={styles.card}>
          <div style={styles.row}>
            <h3 style={styles.title}>Revision {selected.revisionNumber}</h3>
            <span
              style={
                selected.status === 'finalized' || selected.status === 'superseded'
                  ? styles.badgeFinalized
                  : styles.badge
              }
            >
              {selected.status}
            </span>
          </div>
          <div style={styles.mono}>contentHash: {selected.contentHash}</div>

          {isFinalized && (
            <div style={styles.warning}>
              This revision is immutable. Editing is disabled. (Server-enforced.)
            </div>
          )}

          <fieldset
            style={{ border: `1px solid var(--border, #e3e6ea)`, borderRadius: 6, padding: 12 }}
            disabled={isFinalized}
          >
            <legend style={styles.label}>Estimate inputs (editable in draft)</legend>
            <div style={styles.row}>
              <label style={styles.label}>
                Currency
                <input
                  style={styles.input}
                  value={currency}
                  onChange={(e) => setCurrency(e.target.value)}
                />
              </label>
              <label style={styles.label}>
                Overhead %
                <input
                  type="number"
                  step="0.01"
                  style={styles.input}
                  value={overheadPct}
                  onChange={(e) => setOverheadPct(Number(e.target.value))}
                />
              </label>
              <label style={styles.label}>
                Contingency %
                <input
                  type="number"
                  step="0.01"
                  style={styles.input}
                  value={contingencyPct}
                  onChange={(e) => setContingencyPct(Number(e.target.value))}
                />
              </label>
              <label style={styles.label}>
                Profit mode
                <select
                  style={styles.input}
                  value={profitMode}
                  onChange={(e) => setProfitMode(e.target.value as 'markup' | 'margin')}
                >
                  <option value="markup">markup</option>
                  <option value="margin">margin</option>
                </select>
              </label>
              <label style={styles.label}>
                Profit ratio
                <input
                  type="number"
                  step="0.01"
                  style={styles.input}
                  value={profitRatio}
                  onChange={(e) => setProfitRatio(Number(e.target.value))}
                />
              </label>
            </div>
            <div style={styles.row}>
              <label style={styles.label}>
                Line description
                <input
                  style={styles.input}
                  value={lineDesc}
                  onChange={(e) => setLineDesc(e.target.value)}
                />
              </label>
              <label style={styles.label}>
                Quantity
                <input
                  type="number"
                  style={styles.input}
                  value={lineQty}
                  onChange={(e) => setLineQty(Number(e.target.value))}
                />
              </label>
              <label style={styles.label}>
                Unit
                <input
                  style={styles.input}
                  value={lineUnit}
                  onChange={(e) => setLineUnit(e.target.value)}
                />
              </label>
              <label style={styles.label}>
                Rate (minor)
                <input
                  type="number"
                  style={styles.input}
                  value={lineRate}
                  onChange={(e) => setLineRate(Number(e.target.value))}
                />
              </label>
            </div>
            {!isFinalized && (
              <button style={styles.buttonPrimary} onClick={saveDraft}>
                Save draft
              </button>
            )}
          </fieldset>

          {replay && (
            <div style={styles.card}>
              <h4 style={styles.title}>Authoritative totals (from server replay)</h4>
              <div style={styles.muted}>
                hash matches: {replay.contentHashMatches ? '✓' : '✗ MISMATCH'}
              </div>
              <div style={styles.row}>
                <div>
                  <div style={styles.label}>Line cost</div>
                  <div style={styles.value}>{money(replay.totals.totalLineCost)}</div>
                </div>
                <div>
                  <div style={styles.label}>Overhead</div>
                  <div style={styles.value}>{money(replay.totals.overhead)}</div>
                </div>
                <div>
                  <div style={styles.label}>Contingency</div>
                  <div style={styles.value}>{money(replay.totals.contingency)}</div>
                </div>
                <div>
                  <div style={styles.label}>Total cost</div>
                  <div style={styles.value}>{money(replay.totals.totalCost)}</div>
                </div>
                <div>
                  <div style={styles.label}>Profit</div>
                  <div style={styles.value}>{money(replay.totals.profit)}</div>
                </div>
                <div>
                  <div style={styles.label}>Sell price</div>
                  <div style={styles.value}>{money(replay.totals.sellPrice)}</div>
                </div>
                <div>
                  <div style={styles.label}>Gross margin</div>
                  <div style={styles.value}>{(replay.totals.grossMargin * 100).toFixed(2)}%</div>
                </div>
              </div>
            </div>
          )}

          {!isFinalized &&
            selected.status === 'draft' &&
            (confirmingFinalize ? (
              <div style={styles.card}>
                <div style={styles.warning}>Finalize this estimate? It becomes immutable.</div>
                <div style={styles.row}>
                  <button style={styles.buttonPrimary} onClick={finalize}>
                    Confirm finalize
                  </button>
                  <button style={styles.button} onClick={() => setConfirmingFinalize(false)}>
                    Cancel
                  </button>
                </div>
              </div>
            ) : (
              <button style={styles.buttonPrimary} onClick={() => setConfirmingFinalize(true)}>
                Finalize estimate
              </button>
            ))}
        </div>
      )}
    </div>
  )
}
