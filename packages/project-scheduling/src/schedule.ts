import type { CalendarId, DerivedSchedule, ISODateTime, ProjectDocument, Task, TaskId, TaskSchedule, WorkingMinutes } from '@genoffice/project-contracts'
import { validateProjectDocument } from '@genoffice/project-engine'
import { addWorkingTime, resolveCalendar, subtractWorkingTime, workingDuration } from './calendar.js'
import { buildDependencyGraph } from './graph.js'

export interface SchedulingOptions { projectStart?: ISODateTime }

function maxIso(values: Array<ISODateTime | undefined>): ISODateTime | undefined {
  const filtered = values.filter(Boolean) as ISODateTime[]
  return filtered.length ? filtered.reduce((a, b) => (a > b ? a : b)) : undefined
}
function minIso(values: Array<ISODateTime | undefined>): ISODateTime | undefined {
  const filtered = values.filter(Boolean) as ISODateTime[]
  return filtered.length ? filtered.reduce((a, b) => (a < b ? a : b)) : undefined
}
function applyConstraintStart(task: Task, proposed: ISODateTime): ISODateTime {
  if (!task.constraintDate) return proposed
  switch (task.constraintType) {
    case 'startNoEarlierThan': case 'mustStartOn': return proposed < task.constraintDate ? task.constraintDate : proposed
    default: return proposed
  }
}
function applyConstraintFinish(task: Task, proposed: ISODateTime): ISODateTime {
  if (!task.constraintDate) return proposed
  switch (task.constraintType) {
    case 'finishNoEarlierThan': case 'mustFinishOn': return proposed < task.constraintDate ? task.constraintDate : proposed
    default: return proposed
  }
}

export function schedule(document: ProjectDocument, options: SchedulingOptions = {}): DerivedSchedule {
  const validation = validateProjectDocument(document)
  if (!validation.accepted) return { taskSchedules: {}, diagnostics: validation.diagnostics.map((d) => ({ code: d.code, severity: 'error', message: d.message })) }
  const graph = buildDependencyGraph(document)
  const calendars = { calendars: document.calendars }
  const taskById = new Map(document.tasks.map((task) => [task.id, task]))
  const schedules = new Map<TaskId, TaskSchedule>()
  const projectStart = options.projectStart ?? document.properties.startDate

  for (const taskId of graph.topologicalOrder) {
    const task = taskById.get(taskId)!
    const calendarId = task.calendarId ?? document.properties.defaultCalendarId as CalendarId
    const calendar = resolveCalendar(calendars, calendarId)
    let candidateStart: ISODateTime = task.start ?? projectStart
    for (const dependency of graph.predecessors.get(taskId) ?? []) {
      const predecessor = schedules.get(dependency.predecessorId)
      if (!predecessor) continue
      let requiredStart: ISODateTime | undefined
      switch (dependency.type) {
        case 'FS': requiredStart = predecessor.scheduledFinish ? addWorkingTime(calendar, predecessor.scheduledFinish, dependency.lagMinutes as WorkingMinutes) : undefined; break
        case 'SS': requiredStart = predecessor.scheduledStart ? addWorkingTime(calendar, predecessor.scheduledStart, dependency.lagMinutes as WorkingMinutes) : undefined; break
        case 'FF': {
          if (predecessor.scheduledFinish) {
            const requiredFinish = addWorkingTime(calendar, predecessor.scheduledFinish, dependency.lagMinutes as WorkingMinutes)
            requiredStart = subtractWorkingTime(calendar, requiredFinish, task.duration)
          }
          break
        }
        case 'SF': {
          if (predecessor.scheduledStart) {
            const requiredFinish = addWorkingTime(calendar, predecessor.scheduledStart, dependency.lagMinutes as WorkingMinutes)
            requiredStart = subtractWorkingTime(calendar, requiredFinish, task.duration)
          }
          break
        }
      }
      candidateStart = maxIso([candidateStart, requiredStart]) ?? candidateStart
    }
    candidateStart = applyConstraintStart(task, candidateStart)
    const nominalFinish = addWorkingTime(calendar, candidateStart, task.duration)
    const candidateFinish = applyConstraintFinish(task, nominalFinish)
    if (candidateFinish !== nominalFinish) candidateStart = subtractWorkingTime(calendar, candidateFinish, task.duration)
    schedules.set(taskId, {
      taskId,
      earlyStart: candidateStart,
      earlyFinish: candidateFinish,
      totalSlack: 0,
      freeSlack: 0,
      critical: false,
      scheduledStart: candidateStart,
      scheduledFinish: candidateFinish,
      duration: task.duration,
    })
  }

  const leafFinishes = document.tasks.filter((task) => !task.summary).map((task) => schedules.get(task.id)?.scheduledFinish).filter(Boolean) as ISODateTime[]
  const projectFinish = maxIso(leafFinishes)

  for (const task of document.tasks.filter((candidate) => candidate.summary).slice().sort((a, b) => b.outlineLevel - a.outlineLevel || String(b.id).localeCompare(String(a.id)))) {
    const descendants = document.tasks.filter((candidate) => candidate.parentTaskId === task.id)
    const childSchedules = descendants.map((child) => schedules.get(child.id)).filter(Boolean) as TaskSchedule[]
    const start = minIso(childSchedules.map((item) => item.scheduledStart))
    const finish = maxIso(childSchedules.map((item) => item.scheduledFinish))
    if (start && finish) {
      const calendar = resolveCalendar(calendars, (task.calendarId ?? document.properties.defaultCalendarId) as CalendarId)
      schedules.set(task.id, { taskId: task.id, earlyStart: start, earlyFinish: finish, scheduledStart: start, scheduledFinish: finish, duration: workingDuration(calendar, start, finish), totalSlack: 0, freeSlack: 0, critical: false })
    }
  }

  if (projectFinish) {
    for (const taskId of graph.topologicalOrder.slice().reverse()) {
      const task = taskById.get(taskId)!
      const current = schedules.get(taskId)!
      const calendar = resolveCalendar(calendars, (task.calendarId ?? document.properties.defaultCalendarId) as CalendarId)
      let lateStart: ISODateTime
      if (!(graph.successors.get(taskId)?.length)) {
        lateStart = subtractWorkingTime(calendar, projectFinish, task.duration)
      } else {
        const candidates: ISODateTime[] = []
        for (const dependency of graph.successors.get(taskId) ?? []) {
          const successor = schedules.get(dependency.successorId)!
          if (dependency.type === 'SS' || dependency.type === 'SF') {
            if (successor.lateStart) candidates.push(subtractWorkingTime(calendar, successor.lateStart, dependency.lagMinutes as WorkingMinutes))
          } else if (successor.lateFinish) {
            const latestFinish = subtractWorkingTime(calendar, successor.lateFinish, dependency.lagMinutes as WorkingMinutes)
            candidates.push(subtractWorkingTime(calendar, latestFinish, task.duration))
          }
        }
        lateStart = minIso(candidates) ?? subtractWorkingTime(calendar, projectFinish, task.duration)
      }
      const lateFinish = addWorkingTime(calendar, lateStart, task.duration)
      const totalSlack = current.earlyStart ? workingDuration(calendar, current.earlyStart, lateStart) : 0
      const successorStarts = (graph.successors.get(taskId) ?? []).map((dependency) => schedules.get(dependency.successorId)?.earlyStart).filter(Boolean) as ISODateTime[]
      const freeSlack = successorStarts.length && current.earlyFinish ? Math.max(0, workingDuration(calendar, current.earlyFinish, minIso(successorStarts)!)) : totalSlack
      schedules.set(taskId, { ...current, lateStart, lateFinish, totalSlack, freeSlack, critical: totalSlack === 0 })
    }
  }

  const taskSchedules = Object.fromEntries(document.tasks.map((task) => [task.id, schedules.get(task.id)!])) as Record<TaskId, TaskSchedule>
  return { taskSchedules, projectStart, projectFinish, diagnostics: [] }
}
