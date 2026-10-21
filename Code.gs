/** Wolk Pathfinder 3.3 — Google Apps Script receiver.
 * Deploy as Web app: Execute as Me; access Anyone.
 * IMPORTANT: public endpoint. Apply rate limits and monitoring before wide distribution.
 */
const SHEET_ID = 'PASTE_GOOGLE_SHEET_ID_HERE';
const TAB = 'Submissions';
const HEADERS = ['Received at','Submission ID','Name','Email','Game top pathway','Game pathways JSON','Game contributions JSON','Game actions JSON','Self interests','Preferred responsibility','Strengths','Growth interests','Availability','Other notes','Profile consent','Political consent','Raw JSON'];
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents || e.postData.contents.length > 15000) throw Error('Invalid payload');
    const d = JSON.parse(e.postData.contents);
    if (d.schemaVersion !== '3.3' || d.consent?.profile !== true) throw Error('Consent or version missing');
    const c = d.contact || {}, self = d.self || {}, game = d.game || {};
    if (!String(c.name||'').trim() && !String(c.email||'').trim()) throw Error('Name or email required');
    const book = SpreadsheetApp.openById(SHEET_ID);
    const lock = LockService.getScriptLock();lock.waitLock(10000);
    try {
      let sheet=book.getSheetByName(TAB);if(!sheet)sheet=book.insertSheet(TAB);
      if(sheet.getLastRow()===0)sheet.appendRow(HEADERS);
      const safe=x=>{let s=String(x??'').slice(0,4000);return /^[=+\-@\t\r]/.test(s)?"'"+s:s};
      const json=x=>safe(JSON.stringify(x||[]));
      const id=Utilities.getUuid();
      sheet.appendRow([new Date(),id,safe(c.name),safe(c.email),safe(game.pathways?.[0]?.name),json(game.pathways),json(game.contributions),json(game.actions),safe((self.interests||[]).join('; ')),safe(self.responsibility),safe(self.strengths),safe(self.growth),safe(self.availability),safe(self.note),true,d.consent?.political===true,safe(JSON.stringify(d))]);
      return reply({ok:true,id});
    } finally {lock.releaseLock()}
  } catch(err){return reply({ok:false,error:String(err.message||err)})}
}
function reply(data){return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON)}
