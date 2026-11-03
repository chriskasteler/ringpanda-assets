const NOTIFY = 'chris.kasteler@gmail.com';
const FIELDS = ['name','business','phone','email','best_time','note'];

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  let ss;
  if (id) { ss = SpreadsheetApp.openById(id); }
  else {
    ss = SpreadsheetApp.create('Ring Panda setup call requests');
    ss.getSheets()[0].appendRow(['Received','Name','Business','Phone','Email','Best time to call','Note','Page']);
    props.setProperty('SHEET_ID', ss.getId());
  }
  return ss.getSheets()[0];
}

function setup() { const s = getSheet_(); Logger.log(s.getParent().getUrl()); }

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { d = e.parameter || {}; }
  if (d.website) return out_({ok:true}); // honeypot
  const v = FIELDS.map(f => String(d[f] || '').slice(0, 1000));
  if (!v[0] || (!v[2] && !v[3])) return out_({ok:false});
  const sheet = getSheet_();
  sheet.appendRow([new Date()].concat(v).concat([String(d.page || '').slice(0,200)]));
  const body = 'New Ring Panda setup call request\n\n' +
    'Name: ' + v[0] + '\nBusiness: ' + v[1] + '\nPhone: ' + v[2] + '\nEmail: ' + v[3] +
    '\nBest time to call: ' + v[4] + '\nNote: ' + v[5] + '\n\nAll requests: ' + sheet.getParent().getUrl();
  const opts = { name: 'Ring Panda website' };
  if (v[3] && v[3].indexOf('@') > 0) opts.replyTo = v[3];
  MailApp.sendEmail(NOTIFY, 'Ring Panda setup call request: ' + (v[1] || v[0]), body, opts);
  return out_({ok:true});
}

function doGet() { return out_({ok:true, service:'ringpanda-contact'}); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
