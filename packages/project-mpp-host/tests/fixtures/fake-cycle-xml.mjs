// Fake sidecar: CANONICAL REJECTION — writes a cycle-dependency MSPDI that
// the engine must reject (atomic failure at stage 'canonical').
import { writeFileSync } from 'node:fs'
const [, , , outputPath, requestId] = process.argv
const task = (uid, pred) => `    <Task>
      <UID>${uid}</UID><ID>${uid}</ID><Name>T${uid}</Name>
      <OutlineNumber>${uid}</OutlineNumber><OutlineLevel>1</OutlineLevel>
      <Summary>0</Summary><Milestone>0</Milestone>
      <Type>0</Type><Duration>PT8H0M0S</Duration><Work>PT0H0M0S</Work>
      <PercentComplete>0</PercentComplete><Priority>500</Priority>
      ${pred}
    </Task>`
const link = (predUid) =>
  `      <PredecessorLink><PredecessorUID>${predUid}</PredecessorUID><Type>1</Type><LinkLag>0</LinkLag><LinkLagFormat>1</LinkLagFormat></PredecessorLink>`
writeFileSync(
  outputPath,
  `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <SaveVersion>16</SaveVersion><Name>Cycle</Name><ScheduleFromStart>true</ScheduleFromStart>
  <StartDate>2026-08-03T09:00:00</StartDate>
  <Calendars>
    <Calendar><UID>1</UID><Name>Standard</Name><IsBaseCalendar>true</IsBaseCalendar><IsBaseCalendarDefault>true</IsBaseCalendarDefault>
      <WeekDays>
        <WeekDay><DayType>1</DayType><DayWorking>false</DayWorking></WeekDay>
        <WeekDay><DayType>2</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
        <WeekDay><DayType>3</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
        <WeekDay><DayType>4</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
        <WeekDay><DayType>5</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
        <WeekDay><DayType>6</DayType><DayWorking>true</DayWorking><WorkingTimes><WorkingTime><FromTime>09:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
        <WeekDay><DayType>7</DayType><DayWorking>false</DayWorking></WeekDay>
      </WeekDays>
      <Exceptions />
    </Calendar>
  </Calendars>
  <Tasks>
${task(1, link(2))}
${task(2, link(1))}
  </Tasks>
  <Resources></Resources><Assignments></Assignments>
</Project>
`,
)
console.log(
  JSON.stringify({
    version: 1,
    requestId,
    ok: true,
    counts: { tasks: 2, resources: 0, calendars: 1, predecessorLinks: 2, assignments: 0 },
  }),
)
