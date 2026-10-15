# Seeds a temporary new-version project with a defect, a note (with photo and pin), and a temporary old-version project.
import json,urllib.request,io,uuid,sys
from PIL import Image, ImageDraw
B=sys.argv[1] if len(sys.argv)>1 else "http://localhost:5058"; C="2026-10-15T00:00:00.000Z"
def call(m,p,b=None):
    r=urllib.request.Request(B+p,data=json.dumps(b).encode() if b is not None else None,headers={"Content-Type":"application/json"},method=m)
    return json.load(urllib.request.urlopen(r,timeout=60))
def upload(p, field, fname, data, mime, extra={}):
    bd=uuid.uuid4().hex; body=b""
    for k,v in extra.items(): body+=f"--{bd}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n".encode()
    body+=f"--{bd}\r\nContent-Disposition: form-data; name=\"{field}\"; filename=\"{fname}\"\r\nContent-Type: {mime}\r\n\r\n".encode()+data+f"\r\n--{bd}--\r\n".encode()
    r=urllib.request.Request(B+p,data=body,headers={"Content-Type":f"multipart/form-data; boundary={bd}"},method="POST")
    return json.load(urllib.request.urlopen(r,timeout=60))
def img(fmt,color,size):
    b=io.BytesIO(); im=Image.new("RGB",size,color); ImageDraw.Draw(im).rectangle([100,100,400,300],outline=(0,0,0),width=8); im.save(b,fmt); return b.getvalue()
pid=call("POST","/api/projects",{"name":"ZZ Notes check (temporary)","address":"1 St","client":"C","inspector":"LA","createdAt":C,"afcReference":"AFC-26001"})['id']
call("PUT",f"/api/projects/{pid}/categories",[{"code":"RR","label":"Rectify"},{"code":"PN","label":"Project Note","isDefault":True}])
rid=call("POST",f"/api/projects/{pid}/reports",{"inspectionNumber":"12","inspectionDate":"2026-10-15","createdAt":C})['id']
base={"reportId":rid,"dateOpened":"2026-10-15","verificationMethod":"","verificationPerson":"","status":"open","createdAt":C,"elevationCode":"E","dropCode":"04","levelCode":"07","workTypeCode":"GW","seqNumber":"01"}
d=call("POST",f"/api/projects/{pid}/defects",{**base,"uidPrefix":"E-04-07-GW","uidOverride":"E-04-07-GW-01","recordType":"defect","comment":"Cracked render","actionRequired":"Contractor to repair render","categoryCode":"RR","assignedTo":"Contractor"})
note=call("POST",f"/api/projects/{pid}/defects",{**base,"uidPrefix":"E-04-07-GW","uidOverride":"E-04-07-GW-02","recordType":"observation","actionNeeded":False,"comment":"Sealant joints in good condition","actionRequired":""})
upload(f"/api/defects/{note['id']}/photos","photo","n.jpg",img("JPEG",(120,160,200),(800,600)),"image/jpeg",{"slot":"wip1"})
eid=upload(f"/api/projects/{pid}/elevations","file","e.png",img("PNG",(245,245,245),(1200,800)),"image/png",{"name":"East"})['id']
call("POST",f"/api/elevations/{eid}/markers",{"defectId":note['id'],"defectUid":note['uid'],"status":"open","xPercent":30,"yPercent":40})
call("POST",f"/api/elevations/{eid}/markers",{"defectId":d['id'],"defectUid":d['uid'],"status":"open","xPercent":60,"yPercent":40})
call("POST",f"/api/reports/{rid}/action-summaries/regenerate")
opid=call("POST","/api/projects",{"name":"ZZ Old-wording check (temporary)","address":"1 St","client":"C","inspector":"LA","createdAt":C,"featureVersion":1})['id']
orid=call("POST",f"/api/projects/{opid}/reports",{"inspectionNumber":"03","inspectionDate":"2026-10-15","createdAt":C})['id']
out=dict(P=pid,R=rid,D=d['id'],N=note['id'],NUID=note['uid'],E=eid,OP=opid,OR=orid)
json.dump(out,open('/tmp/wp5seed.json','w')); print(out)
