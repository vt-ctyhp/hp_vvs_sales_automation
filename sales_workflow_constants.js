/**
 * Sales workflow constants: sheet names, statuses, task types, and tab headers.
 */

var SW_SHEETS = {
  MASTER: '00_Master Appointments',
  TASKS: '_SalesTaskQueue',
  LOG: '_SalesTaskLog',
  CONFIG: '_SalesWorkflowConfig',
  TEMPLATES: '_SalesWorkflowTemplates',
  USERS: '_SalesWorkflowUsers',
  ROSTER: '10_Roster_Schedule',
  SCHEDULE_CHANGES: 'Schedule Changes',
  DROPDOWN: 'Dropdown'
};
var SW_STATUSES = {
  PENDING: 'Pending',
  COMPLETED: 'Completed',
  BLOCKED: 'Blocked',
  SNOOZED: 'Snoozed'
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
  FINAL: 'SEND_FINAL_RECAP',
  POST_CONSULT_STATUS: 'POST_CONSULT_CLIENT_STATUS',
  START_3D: 'START_3D_DESIGN',
  RECORD_3D_DEADLINE: 'RECORD_3D_DEADLINE',
  REQUEST_WAX: 'REQUEST_WAX_PRINT',
  UPDATE_WAX: 'UPDATE_WAX_REQUEST',
  DIAMOND_PROPOSE: 'PROPOSE_DIAMONDS',
  DIAMOND_QUOTE: 'PREPARE_DV_QUOTATION',
  DIAMOND_ORDER: 'ORDER_DIAMONDS',
  DIAMOND_TRACK: 'TRACK_DIAMONDS',
  DIAMOND_DELIVERY: 'CONFIRM_DIAMOND_DELIVERY',
  DIAMOND_DECISIONS: 'RECORD_DIAMOND_DECISIONS',
  DIAMOND_RETURN: 'RETURN_DIAMONDS',
  DIAMOND_ORDER_ACK_REP: 'ACK_DIAMONDS_ORDERED_ASSIGNED_REP',
  DIAMOND_ORDER_ACK_JOC: 'ACK_DIAMONDS_ORDERED_JOC',
  DIAMOND_ETA_REP: 'REVIEW_DIAMOND_ETA_ASSIGNED_REP',
  DIAMOND_ETA_JOC: 'REVIEW_DIAMOND_ETA_JOC'
};
var SW_OWNER_ROLES = {
  SYSTEM: 'System',
  SALES_REP: 'SALES_REP',
  JOC: 'JOC',
  DIAMOND_ORDER_ADMIN: 'DIAMOND_ORDER_ADMIN',
  DIAMOND_ORDER_ASSISTANT: 'DIAMOND_ORDER_ASSISTANT'
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
  'Primary Action',
  'Snooze Until',
  'Snooze Reason',
  'Snoozed By',
  'Snoozed At'
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
  'Primary Action',
  'Snooze Until',
  'Snooze Reason'
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
var SW_AUTH_USER_HEADERS = [
  'Email',
  'Name',
  'Roles',
  'Active?',
  'Password Salt',
  'Password Hash',
  'Temporary Password?',
  'Last Login At',
  'Notes'
];
