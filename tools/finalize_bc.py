import json,sys
"""แปลง data.js ที่สร้างจากไฟล์ Block Course ให้เป็น data-bc.js (window.TT_BC)
ใช้: python tools/finalize_bc.py data-bc.js"""
p=sys.argv[1]; s=open(p).read(); T=json.loads(s[s.index('{'):s.rindex('}')+1])
T['weeks']='16–18'
T['note']='ปวส.1 เรียนต่อหลังจบภาคปกติ (ปวส. 15 สัปดาห์)'
for t in T.get('teachers',[]): t['full']=t['full'].replace('นาวสาว','นางสาว')
for c in T['classes']: c['advisor']=c['advisor'].replace('นาวสาว','นางสาว')
open(p,'w').write('window.TT_BC='+json.dumps(T,ensure_ascii=False,separators=(',',':'))+';\n')
print(len(T['classes']),'classes,',len(T.get('teachers',[])),'teachers')
