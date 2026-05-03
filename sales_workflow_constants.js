/**
 * Sales workflow constants: sheet names, statuses, task types, and tab headers.
 */

var SW_SHEETS = {
  MASTER: '00_Master Appointments',
  TASKS: '_SalesTaskQueue',
  LOG: '_SalesTaskLog',
  CONFIG: '_SalesWorkflowConfig',
  TEMPLATES: '_SalesWorkflowTemplates',
  ROSTER: '10_Roster_Schedule',
  SCHEDULE_CHANGES: 'Schedule Changes',
  DROPDOWN: 'Dropdown'
};
var SW_STATUSES = {
  PENDING: 'Pending',
  COMPLETED: 'Completed',
  BLOCKED: 'Blocked'
};
var SW_TASKS = {
  ASSIGN: 'ASSIGN_APPOINTMENT',
  WELCOME: 'SEND_WELCOME',
  HYBRID: 'SEND_HYBRID_WELCOME',
  MAP: 'SEND_MAP_INSTRUCTIONS',
  REVIEW: 'REVIEW_APPOINTMENT',
  CHECKLIST: 'APPOINTMENT_DAY_CHECKLIST',
  PROCESS: 'PROCESS_APPOINTMENT_DATA',
  APPROVE: 'APPROVE_RECAP_MESSAGE',
  FINAL: 'SEND_FINAL_RECAP'
};
var SW_TASK_HEADERS = [
  'TaskID',
  'RootApptID',
  'APPT_ID',
  'Customer Name',
  'Brand',
  'Visit Date',
  'Visit Time',
  'Visit Type',
  'Lifecycle Stage',
  'Task Type',
  'Task Title',
  'Owner Role',
  'Intended Owner',
  'Intended Owner Email',
  'Current Owner',
  'Current Owner Email',
  'Coverage Reason',
  'Due At',
  'Status',
  'Dependency TaskID',
  'Created At',
  'Updated At',
  'Completed By',
  'Completed By Email',
  'Completed At',
  'Claimed By',
  'Claimed At',
  'Last Event',
  'Payload JSON',
  'Template Key',
  'Instructions',
  'Primary Action'
];
var SW_TASK_LIST_HEADERS = [
  'TaskID',
  'RootApptID',
  'APPT_ID',
  'Customer Name',
  'Brand',
  'Visit Date',
  'Visit Time',
  'Visit Type',
  'Lifecycle Stage',
  'Task Type',
  'Task Title',
  'Owner Role',
  'Intended Owner',
  'Intended Owner Email',
  'Current Owner',
  'Current Owner Email',
  'Coverage Reason',
  'Due At',
  'Status',
  'Primary Action'
];
var SW_LOG_HEADERS = [
  'Event At',
  'Event Type',
  'TaskID',
  'RootApptID',
  'APPT_ID',
  'Task Type',
  'Actor Name',
  'Actor Email',
  'From Owner',
  'To Owner',
  'Status',
  'Details JSON'
];
var SW_CONFIG_HEADERS = [
  'Section',
  'Key',
  'Value',
  'Role',
  'Name',
  'Email',
  'Active?',
  'Priority',
  'Notes'
];
var SW_TEMPLATE_HEADERS = [
  'Task Type',
  'Task Title',
  'Instructions',
  'Template',
  'Attachment Label',
  'Attachment URL',
  'Checklist JSON',
  'Primary Action'
];
