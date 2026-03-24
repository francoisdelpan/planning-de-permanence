const PERMANENCE_CODE = 'P';
const PERMANENCE_TITLE = 'Permanence';
const PERMANENCE_START_HOUR = 12;
const PERMANENCE_END_HOUR = 20;

const TEMPLATE_SHEET = 'TRAME';
const NEW_SHEET_PREFIX = ''; // vide pour garder juste le n° de semaine

const WEEK_LABEL_FORMAT = 'w yyyy'; // format pour nommer les semaines
const MAX_WEEKS = 8; // sécurité
const START_CELL_DATE = 'D7';
const START_CELL_WEEK = 'D2';
const LINE_DATE = 7;
const COL_TRAME_USER = 2;
const LINE_TRAME_USER_START = 8;

const DB_SHEET = 'DATABASE';
const COL_DATECODE = 0;
const COL_NAME = 1;
const COL_CALENDAR = 2;
const COL_OLDVALUE = 3;
const COL_NEWVALUE = 4;
const COL_STATUS = 5;

const ADMIN_EMAIL = 'gaelle_botineau@franchise.carrefour.com';
const PLANNING_URL = 'https://docs.google.com/spreadsheets/d/193WPmJIK3RDlXJO7yj576Dk1CiqvYqInCx1QGyxjbIU/edit?gid=0#gid=0';
const DAYS_FR = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];

function testCalendarAuth() {
  const cal = CalendarApp.getCalendarById("francois_pannier@franchise.carrefour.com");
  cal.createAllDayEvent("Test Auth", new Date());
}