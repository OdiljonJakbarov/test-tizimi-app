from flask import Flask, render_template, request, jsonify, send_file, session
import os, json, random, time, sqlite3
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import uuid

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'test_app_secret_2024_xyz')

BASE_DIR = os.environ.get('RAILWAY_VOLUME_MOUNT_PATH', os.path.dirname(os.path.abspath(__file__)))
TESTS_DIR = os.path.join(BASE_DIR, 'tests')
RESULTS_DIR = os.path.join(BASE_DIR, 'results')
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')
DB_FILE = os.path.join(BASE_DIR, 'results.db')

DEFAULT_CONFIG = {"time_limit": 30, "question_count": 10, "admin_password": "admin123"}

def init_db():
    os.makedirs(TESTS_DIR, exist_ok=True)
    os.makedirs(RESULTS_DIR, exist_ok=True)
    con = sqlite3.connect(DB_FILE)
    con.execute('''CREATE TABLE IF NOT EXISTS results (
        id TEXT PRIMARY KEY, fio TEXT, grp TEXT, category TEXT,
        total INTEGER, correct INTEGER, wrong INTEGER,
        percentage REAL, date TEXT, answers TEXT)''')
    con.commit(); con.close()

def db_save(r):
    con = sqlite3.connect(DB_FILE)
    con.execute('''INSERT OR REPLACE INTO results
        (id,fio,grp,category,total,correct,wrong,percentage,date,answers)
        VALUES (?,?,?,?,?,?,?,?,?,?)''',
        (r['id'],r['fio'],r['group'],r['category'],r['total'],r['correct'],
         r['wrong'],r['percentage'],r['date'],json.dumps(r['answers'],ensure_ascii=False)))
    con.commit(); con.close()

def db_load(category=None):
    con = sqlite3.connect(DB_FILE)
    con.row_factory = sqlite3.Row
    if category and category != 'all':
        rows = con.execute('SELECT * FROM results WHERE category=? ORDER BY date DESC',(category,)).fetchall()
    else:
        rows = con.execute('SELECT * FROM results ORDER BY date DESC').fetchall()
    con.close()
    results=[]
    for row in rows:
        d=dict(row); d['group']=d.pop('grp'); d['answers']=json.loads(d['answers']); results.append(d)
    return results

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE,'r',encoding='utf-8') as f: return json.load(f)
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    with open(CONFIG_FILE,'w',encoding='utf-8') as f: json.dump(cfg,f,ensure_ascii=False,indent=2)

def get_test_categories():
    categories={}
    os.makedirs(TESTS_DIR,exist_ok=True)
    for fname in sorted(os.listdir(TESTS_DIR)):
        if fname.endswith('.xlsx'): categories[fname[:-5]]=os.path.join(TESTS_DIR,fname)
    return categories

def load_questions(filepath):
    df=pd.read_excel(filepath,header=None)
    questions=[]
    for _,row in df.iterrows():
        vals=[str(v).strip() for v in row.tolist() if str(v).strip() not in ('nan','')]
        if len(vals)>=2: questions.append({'question':vals[0],'correct':vals[1],'options':vals[1:]})
    return questions

def get_on_categories(categories):
    on1=[k for k in categories if '1-' in k and ('ON' in k.upper() or 'ОН' in k)]
    on2=[k for k in categories if '2-' in k and ('ON' in k.upper() or 'ОН' in k)]
    return on1,on2

@app.route('/')
def index():
    config=load_config(); categories=get_test_categories()
    on1,on2=get_on_categories(categories)
    return render_template('index.html',categories=list(categories.keys()),has_yn=bool(on1 and on2),config=config)

@app.route('/start_test',methods=['POST'])
def start_test():
    data=request.json
    fio=data.get('fio','').strip(); group=data.get('group','').strip(); category=data.get('category','')
    if not fio or not group: return jsonify({'error':'FIO va guruh kiritilishi shart!'}),400
    config=load_config(); categories=get_test_categories(); questions=[]
    if category=='YaN':
        on1,on2=get_on_categories(categories)
        if not on1 or not on2: return jsonify({'error':'YaN uchun 1-ON va 2-ON fayllari zarur!'}),400
        q_count=config.get('question_count',10); half=q_count//2
        q1=load_questions(categories[on1[0]]); q2=load_questions(categories[on2[0]])
        random.shuffle(q1); random.shuffle(q2)
        questions=q1[:half]+q2[q_count-half:]
    else:
        if category not in categories: return jsonify({'error':'Kategoriya topilmadi!'}),400
        all_q=load_questions(categories[category]); random.shuffle(all_q)
        questions=all_q[:config.get('question_count',10)]
    for q in questions:
        opts=q['options'][:]; random.shuffle(opts); q['shuffled_options']=opts
    session['test']={'id':str(uuid.uuid4()),'fio':fio,'group':group,'category':category,
        'questions':questions,'start_time':time.time(),'time_limit':config.get('time_limit',30)*60,'answers':[]}
    return jsonify({'success':True,'total':len(questions)})

@app.route('/test')
def test_page():
    if 'test' not in session: return render_template('error.html',msg='Test topilmadi. Qaytadan boshlang.')
    t=session['test']
    remaining=max(0,t['time_limit']-(time.time()-t['start_time']))
    return render_template('test.html',fio=t['fio'],group=t['group'],category=t['category'],total=len(t['questions']),time_limit=int(remaining))

@app.route('/get_question/<int:idx>')
def get_question(idx):
    if 'test' not in session: return jsonify({'error':'Session tugagan'}),400
    t=session['test']
    elapsed=time.time()-t['start_time']
    if elapsed>t['time_limit']: return jsonify({'timeout':True})
    if idx>=len(t['questions']): return jsonify({'done':True})
    q=t['questions'][idx]
    return jsonify({'question':q['question'],'options':q['shuffled_options'],'index':idx,'total':len(t['questions']),'remaining':int(t['time_limit']-elapsed)})

@app.route('/submit_answer',methods=['POST'])
def submit_answer():
    if 'test' not in session: return jsonify({'error':'Session tugagan'}),400
    data=request.json; t=session['test']
    q=t['questions'][data['index']]; correct=(data['answer']==q['correct'])
    t['answers'].append({'index':data['index'],'answer':data['answer'],'correct':correct})
    session['test']=t
    return jsonify({'correct':correct,'correct_answer':q['correct']})

@app.route('/finish_test',methods=['POST'])
def finish_test():
    if 'test' not in session: return jsonify({'error':'Session tugagan'}),400
    t=session['test']; total=len(t['questions'])
    correct_count=sum(1 for a in t['answers'] if a['correct'])
    percentage=round(correct_count/total*100,1) if total>0 else 0
    result={'id':t['id'],'fio':t['fio'],'group':t['group'],'category':t['category'],'total':total,
        'correct':correct_count,'wrong':total-correct_count,'percentage':percentage,
        'date':datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'answers':t['answers']}
    db_save(result); session.pop('test',None)
    return jsonify(result)

@app.route('/admin')
def admin():
    config=load_config(); categories=get_test_categories(); results=db_load()
    return render_template('admin.html',config=config,categories=list(categories.keys()),results=results)

@app.route('/admin/save_config',methods=['POST'])
def save_config_route():
    data=request.json; config=load_config()
    if data.get('password')!=config.get('admin_password','admin123'): return jsonify({'error':'Noto\'g\'ri parol!'}),403
    config['time_limit']=int(data.get('time_limit',30)); config['question_count']=int(data.get('question_count',10))
    if data.get('new_password'): config['admin_password']=data['new_password']
    save_config(config); return jsonify({'success':True})

@app.route('/admin/upload_test',methods=['POST'])
def upload_test():
    if 'file' not in request.files: return jsonify({'error':'Fayl yuklanmadi'}),400
    f=request.files['file']; name=request.form.get('name','').strip()
    if load_config().get('admin_password')!=request.form.get('password',''): return jsonify({'error':'Noto\'g\'ri parol!'}),403
    if not name: return jsonify({'error':'Test nomi kiritilmagan'}),400
    if not f.filename.endswith('.xlsx'): return jsonify({'error':'Faqat xlsx fayl qabul qilinadi'}),400
    os.makedirs(TESTS_DIR,exist_ok=True)
    f.save(os.path.join(TESTS_DIR,f"{name}.xlsx"))
    return jsonify({'success':True,'message':f'"{name}" muvaffaqiyatli yuklandi!'})

@app.route('/admin/delete_test',methods=['POST'])
def delete_test():
    data=request.json
    if data.get('password')!=load_config().get('admin_password'): return jsonify({'error':'Noto\'g\'ri parol!'}),403
    path=os.path.join(TESTS_DIR,f"{data.get('name')}.xlsx")
    if os.path.exists(path): os.remove(path); return jsonify({'success':True})
    return jsonify({'error':'Fayl topilmadi'}),404

@app.route('/export_results',methods=['POST'])
def export_results():
    data=request.json
    if data.get('password')!=load_config().get('admin_password'): return jsonify({'error':'Noto\'g\'ri parol!'}),403
    results=db_load(data.get('category','all'))
    if not results: return jsonify({'error':'Natijalar topilmadi'}),404
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Natijalar"
    h_fill=PatternFill("solid",fgColor="1A3A6B"); h_font=Font(bold=True,color="FFFFFF",name="Arial",size=11)
    border=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    even_fill=PatternFill("solid",fgColor="EBF5FB")
    headers=["№","F.I.O.","Guruh","Kategoriya","Savollar soni","To'g'ri javob","Xato javob","Foiz (%)","Sana"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill=h_fill; cell.font=h_font
        cell.alignment=Alignment(horizontal='center',vertical='center'); cell.border=border
    ws.row_dimensions[1].height=28
    for i,r in enumerate(results,1):
        ws.append([i,r['fio'],r['group'],r['category'],r['total'],r['correct'],r['wrong'],r['percentage'],r['date']])
        for cell in ws[ws.max_row]:
            cell.border=border; cell.alignment=Alignment(horizontal='center',vertical='center')
            if i%2==0: cell.fill=even_fill
        pct=ws.cell(row=ws.max_row,column=8)
        if isinstance(pct.value,(int,float)):
            color="1E8449" if pct.value>=85 else ("D4AC0D" if pct.value>=55 else "C0392B")
            pct.font=Font(color=color,bold=True,name="Arial")
    for i,w in enumerate([5,30,15,20,16,14,12,12,22],1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    os.makedirs(RESULTS_DIR,exist_ok=True)
    fname=f"natijalar_{data.get('category','all')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    fpath=os.path.join(RESULTS_DIR,fname)
    wb.save(fpath)
    return send_file(fpath,as_attachment=True,download_name=fname)

with app.app_context():
    init_db()
    if not os.path.exists(CONFIG_FILE): save_config(DEFAULT_CONFIG)

if __name__=='__main__':
    port=int(os.environ.get('PORT',5050))
    app.run(host='0.0.0.0',port=port,debug=False)
