import { registerLeveler } from '@genoffice/project-engine'

export * from './calendar.js'
export * from './graph.js'
export * from './schedule.js'
export * from './baseline.js'
export * from './leveling.js'
export * from './allocation.js'

// PROJECT-013: register the canonical resource leveler with the engine
// package's `LevelResources` command dispatch. The leveler lives in the
// scheduling package (it reads the derived schedule); the engine package is a
// lower architectural layer and cannot statically import the scheduling
// package, so the leveler is injected through the `registerLeveler` slot.
// Importing the scheduling package (which every host that schedules must do)
// populates the slot so the engine's `applyProjectCommand(LevelResources)`
// dispatch can call the leveler without a circular static import.
import { levelResources } from './leveling.js'
registerLeveler(levelResources)
