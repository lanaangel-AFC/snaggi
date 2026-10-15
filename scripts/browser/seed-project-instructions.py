# Seeds a temporary new-version project (one defect, a PI category to prove it is hidden)
# and a temporary old-version project with one defect. Writes /tmp/wp6seed.json.
import json,urllib.request,sys
B=sys.argv[1] if len(sys.argv)>1 else "http://localhost:5059"; C="2026-10-16T00:00:00.000Z"
def call(m,p,b=None):
    r=urllib.request.Request(B+p,data=json.dumps(b).encode() if b is not None else None,headers={"Content-Type":"application/json"},method=m)
    return json.load(urllib.request.urlopen(r,timeout=60))
def proj(name, fv):
    body={"name":name,"address":"1 St","client":"C","inspector":"LA","createdAt":C,"afcReference":"AFC-26001"}
    if fv: body["featureVersion"]=fv
    pid=call("POST","/api/projects",body)['id']
    call("PUT",f"/api/projects/{pid}/categories",[{"code":"RR","label":"Rectify","isDefault":True},{"code":"PI","label":"Project Instruction"}])
    rid=call("POST",f"/api/projects/{pid}/reports",{"inspectionNumber":"01","inspectionDate":"2026-10-16","createdAt":C})['id']
    call("POST",f"/api/projects/{pid}/defects",{"reportId":rid,"dateOpened":"2026-10-16","verificationMethod":"","verificationPerson":"","status":"open","createdAt":C,"elevationCode":"E","dropCode":"04","levelCode":"07","workTypeCode":"SE","seqNumber":"01","uidPrefix":"E-04-07-SE","uidOverride":"E-04-07-SE-01","recordType":"defect","comment":"Sealant failed","actionRequired":"Contractor to replace sealant","categoryCode":"RR","assignedTo":"Contractor"})
    call("POST",f"/api/reports/{rid}/action-summaries/regenerate")
    return pid,rid
P,R=proj("ZZ Instructions check (temporary)",None); OP,OR=proj("ZZ Old-project check (temporary)",1)
out=dict(P=P,R=R,OP=OP,OR=OR); json.dump(out,open('/tmp/wp6seed.json','w')); print(out)
