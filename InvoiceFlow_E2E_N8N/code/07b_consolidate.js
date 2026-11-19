// SAP invoice endpoint may emit many n8n items; collapse to one execution before requesting TCURR.
const items=$input.all();
return [{json:{receivedSAPHttpItems:items.length, errors:items.filter(x=>x.json?.error).map(x=>String(x.json.error?.message||x.json.error))}}];
