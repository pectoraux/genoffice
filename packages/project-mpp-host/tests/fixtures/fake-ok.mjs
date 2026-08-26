// Fake sidecar: SUCCESS — writes a minimal valid MSPDI to the output path
// and prints a protocol v1 ok frame. argv: [node, script, input, output, requestId]
import { writeFileSync } from 'node:fs'

const [, , _inputPath, outputPath, requestId] = process.argv

const MSPDI = `<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <SaveVersion>16</SaveVersion>
  <Name>Fake Project</Name>
  <ScheduleFromStart>true</ScheduleFromStart>
  <StartDate>2026-08-03T09:00:00</StartDate>
  <Calendars>
    <Calendar>
      <UID>1</UID>
      <Name>Standard</Name>
      <IsBaseCalendar>true</IsBaseCalendar>
      <IsBaseCalendarDefault>true</IsBaseCalendarDefault>
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
    <Task>
      <UID>1</UID><ID>1</ID><Name>Only Task</Name>
      <OutlineNumber>1</OutlineNumber><OutlineLevel>1</OutlineLevel>
      <Summary>0</Summary><Milestone>0</Milestone>
      <Type>0</Type><Duration>PT8H0M0S</Duration><Work>PT0H0M0S</Work>
      <PercentComplete>0</PercentComplete><Priority>500</Priority>
    </Task>
  </Tasks>
  <Resources></Resources>
  <Assignments></Assignments>
</Project>
`

writeFileSync(outputPath, MSPDI)
console.log(
  JSON.stringify({
    version: 1,
    requestId,
    ok: true,
    counts: { tasks: 1, resources: 0, calendars: 1, predecessorLinks: 0, assignments: 0 },
  }),
)
