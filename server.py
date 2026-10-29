from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime, getaddresses
from datetime import datetime, date
from urllib.parse import urlparse
import csv, io, json, os, re, hashlib, mimetypes, tempfile
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

ROOT=Path(__file__).parent
MAX_UPLOAD=25*1024*1024

def json_bytes(obj): return json.dumps(obj,ensure_ascii=False,default=str).encode()
def norm(x): return re.sub(r'\s+',' ',str(x or '')).strip()
def parse_date(v):
    if isinstance(v,datetime): return v
    if isinstance(v,date): return datetime.combine(v,datetime.min.time())
    s=norm(v)
    if not s:return None
    for fmt in ('%Y-%m-%d %H:%M:%S','%Y-%m-%d','%d/%m/%Y','%d/%m/%Y %H:%M','%m/%d/%Y'):
        try:return datetime.strptime(s[:19],fmt)
        except ValueError: pass
    try:return parsedate_to_datetime(s).replace(tzinfo=None)
    except Exception:return None

def decode_upload(field):
    item=field.file; data=item.read(MAX_UPLOAD+1)
    if len(data)>MAX_UPLOAD: raise ValueError('Arquivo excede o limite de 25 MB.')
    name=os.path.basename(field.filename or 'importacao')
    if not name: raise ValueError('Nome do arquivo inválido.')
    return name,data

def parse_rows(name,data):
    ext=Path(name).suffix.lower()
    if ext in ('.xlsx','.xlsm'):
        wb=load_workbook(io.BytesIO(data),read_only=True,data_only=True,keep_vba=ext=='.xlsm')
        out=[]
        for ws in wb.worksheets:
            rows=list(ws.iter_rows(values_only=True))
            if not rows:continue
            header_i=0
            for i,r in enumerate(rows[:15]):
                vals=[norm(v).lower() for v in r]
                if any(v in ('data','date','received','sent','remetente','assunto','subject') for v in vals): header_i=i;break
            headers=[norm(v) or f'Campo {j+1}' for j,v in enumerate(rows[header_i])]
            for r in rows[header_i+1:]:
                if not any(v is not None and norm(v) for v in r):continue
                out.append({headers[i]: (r[i] if i<len(r) else None) for i in range(len(headers))})
        return out
    if ext=='.csv':
        text=data.decode('utf-8-sig',errors='replace')
        try:dialect=csv.Sniffer().sniff(text[:4096],delimiters=',;\t')
        except csv.Error:dialect=csv.excel
        return [dict(r) for r in csv.DictReader(io.StringIO(text),dialect=dialect)]
    raise ValueError('Formato não suportado. Para e-mails, use .eml, .csv ou .xlsx/.xlsm de metadados.')

def first(row,*names):
    lookup={re.sub(r'[^a-z0-9]','',str(k).lower()):v for k,v in row.items()}
    for n in names:
        key=re.sub(r'[^a-z0-9]','',n.lower())
        if key in lookup:return lookup[key]
    return None

def parse_eml(name,data):
    msg=BytesParser(policy=policy.default).parsebytes(data)
    dt=parse_date(msg.get('Date'))
    frm=getaddresses([msg.get('From','')]); to=getaddresses(msg.get_all('To',[])); cc=getaddresses(msg.get_all('Cc',[]))
    return {'data':dt.isoformat(sep=' ') if dt else None,'remetente':frm[0][1] if frm else '',
      'destinatarios':'; '.join(a for _,a in to+cc),'assunto':str(msg.get('Subject','')),
      'direcao':'Não identificada','cliente':'','conversa':msg.get('Message-ID',''),'id':msg.get('Message-ID',''),'fonte':name}

def email_normalize(rows, source):
    out=[]; seen=set(); duplicates=0
    for r in rows:
        dt=parse_date(first(r,'data','date','sent','sent date','received','received date','data/hora','timestamp'))
        sender=norm(first(r,'remetente','from','sender','email do remetente'))
        recips=norm(first(r,'destinatarios','destinatários','to','recipient','recipients','cc'))
        subject=norm(first(r,'assunto','subject','title'))
        direction=norm(first(r,'direcao','direção','direction')) or 'Não identificada'
        client=norm(first(r,'cliente','operadora','client','operator','cliente/operadora'))
        thread=norm(first(r,'conversa','thread','conversation id','conversationid','thread id'))
        msgid=norm(first(r,'message id','message-id','internet message id','id'))
        key=msgid or hashlib.sha256('|'.join([str(dt),sender,recips,subject]).encode()).hexdigest()
        if key in seen:duplicates+=1;continue
        seen.add(key)
        out.append({'data':dt.isoformat(sep=' ') if dt else None,'remetente':sender,'destinatarios':recips,'assunto':subject,'direcao':direction,'cliente':client,'conversa':thread,'id':msgid,'fonte':source})
    return out,duplicates

def matrix_parse(name,data):
    ext=Path(name).suffix.lower()
    if ext not in ('.xlsx','.xlsm'):raise ValueError('Matriz: selecione um arquivo .xlsx ou .xlsm não criptografado.')
    wb=load_workbook(io.BytesIO(data),read_only=True,data_only=True,keep_vba=ext=='.xlsm')
    sheets=[]; records=[]; ignored=0; all_codes=set(); no_code=0
    for ws in wb.worksheets:
        rows=list(ws.iter_rows(values_only=True)); header_i=None
        for i,row in enumerate(rows[:20]):
            cells=[norm(v).upper() for v in row]
            if 'OPERADORA' in cells and ('COD' in cells or 'CÓD' in cells):header_i=i;break
        if header_i is None:
            sheets.append({'aba':ws.title,'linhas':max(0,len(rows)-1),'lidas':0,'status':'Cabeçalho operacional não localizado'});continue
        headers=[norm(x) for x in rows[header_i]]
        count=0
        for row in rows[header_i+1:]:
            d={headers[i]: (row[i] if i<len(row) else None) for i in range(len(headers)) if headers[i]}
            op=norm(first(d,'OPERADORA')); client=norm(first(d,'CLIENTE')); code=norm(first(d,'COD','CÓD')); opid=norm(first(d,'ID'))
            if not any((op,client,code,opid)):ignored+=1;continue
            if code:all_codes.add(code)
            else:no_code+=1
            records.append({'cliente':client,'operadora':op,'codigo':code,'id_operadora':opid,'aba':ws.title,'categoria':norm(first(d,'CATEGORIA')),'prioridade':norm(first(d,'PRIORIDADE')),'status':norm(first(d,'STATUS')),'data_contrato':str(first(d,'DATA DO CONTRATO') or ''),'dados':d})
            count+=1
        sheets.append({'aba':ws.title,'linhas':max(0,len(rows)-header_i-1),'lidas':count,'status':'Cabeçalho localizado'})
    # Distinct records use stable matrix code and operator ID where available; keep unique row if neither exists.
    unique={}; dups=0
    for r in records:
        k=(r['codigo'],r['id_operadora'],r['operadora']) if r['codigo'] or r['id_operadora'] else ('sem-codigo',r['aba'],len(unique))
        if k in unique:dups+=1
        else:unique[k]=r
    return {'arquivo':name,'hash':hashlib.sha256(data).hexdigest(),'lido_em':datetime.now().astimezone().isoformat(timespec='seconds'),'abas':sheets,'registros':list(unique.values()),'total_linhas':len(records)+ignored,'registros_unicos':len(unique),'codigos_unicos':len(all_codes),'sem_codigo':no_code,'duplicidades':dups,'ignoradas':ignored}

def make_xlsx(payload):
    wb=Workbook(); ws=wb.active; ws.title='Resumo Executivo'
    meta=payload.get('meta',{}); emails=payload.get('emails',[]); matrix=payload.get('matrix',{})
    ws.append(['Duarte Reports — Relatório operacional']);ws.append(['Período',meta.get('period','Não informado')]);ws.append(['Gerado em',datetime.now().astimezone().isoformat(timespec='seconds')]);ws.append(['Responsável',meta.get('responsible') or 'Não informado']);ws.append(['Fonte',meta.get('source') or 'Importação local']);ws.append(['E-mails únicos',len(emails)]);ws.append(['Clientes/operadoras identificados',len({e.get('cliente') for e in emails if e.get('cliente')})]);ws.append(['Duplicidades removidas',payload.get('duplicates',0)]);ws.append(['Observação','Correspondência de cliente não inferida; valores ausentes preservados.'])
    detail=wb.create_sheet('Detalhamento de E-mails'); detail.append(['Data/hora','Remetente','Destinatários','Assunto','Direção','Cliente/operadora','Conversa','ID mensagem','Fonte'])
    for e in emails: detail.append([e.get(k) for k in ['data','remetente','destinatarios','assunto','direcao','cliente','conversa','id','fonte']])
    clients=wb.create_sheet('Consolidado por Cliente'); clients.append(['Cliente/operadora','Mensagens','Último contato'])
    grouped={}
    for e in emails:
        c=e.get('cliente') or 'Não identificado'; d=e.get('data'); q=grouped.setdefault(c,{'count':0,'last':''});q['count']+=1;q['last']=max(q['last'],d or '')
    for c,v in sorted(grouped.items()):clients.append([c,v['count'],v['last'] or None])
    rev=wb.create_sheet('Matriz e Revisão'); rev.append(['Cliente','Operadora','COD','ID operadora','Aba','Categoria','Prioridade','Status','Data contrato'])
    for r in matrix.get('registros',[]):rev.append([r.get('cliente'),r.get('operadora'),r.get('codigo'),r.get('id_operadora'),r.get('aba'),r.get('categoria'),r.get('prioridade'),r.get('status'),r.get('data_contrato')])
    method=wb.create_sheet('Metodologia');
    for row in [['Item','Descrição'],['Fonte de e-mails',meta.get('source') or 'Importação local'],['Período',meta.get('period','Não informado')],['Contagem','Mensagem única por Message-ID ou hash de data/remetente/destinatários/assunto. Conversas não somadas separadamente.'],['Vínculo cliente','Apenas campo explícito na fonte; sem inferência por destinatário ou assunto.'],['Limitação','Sem OAuth/Graph; sem persistência; associações manuais devem ser revisadas.'],['Matriz',matrix.get('arquivo','Não importada')],['Regras semáforo','Limites provisórios definidos pelo usuário na interface; não são regra oficial.']]:method.append(row)
    for sh in wb.worksheets:
        sh.freeze_panes='A2';sh.auto_filter.ref=sh.dimensions
        for c in sh[1]:c.fill=PatternFill('solid',fgColor='001E57');c.font=Font(name='Arial',bold=True,color='FFFFFF')
        for col in sh.columns:
            letter=get_column_letter(col[0].column); width=min(48,max(12,max((len(str(c.value or '')) for c in col[:100]),default=10)+2));sh.column_dimensions[letter].width=width
        for row in sh.iter_rows():
            for c in row:c.font=Font(name='Arial',bold=c.font.bold,color=c.font.color.rgb if c.font.color and c.font.color.type=='rgb' else '001E57',size=10)
    b=io.BytesIO();wb.save(b);return b.getvalue()

def make_pdf(payload):
    meta=payload.get('meta',{}); emails=payload.get('emails',[]); matrix=payload.get('matrix',{}); buf=io.BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=A4,rightMargin=18*mm,leftMargin=18*mm,topMargin=18*mm,bottomMargin=18*mm,title='Duarte Reports — Relatório operacional')
    styles=getSampleStyleSheet();styles.add(ParagraphStyle(name='DRTitle',parent=styles['Title'],textColor=colors.HexColor('#001E57'),fontSize=25,leading=30,spaceAfter=9));styles.add(ParagraphStyle(name='DRH',parent=styles['Heading2'],textColor=colors.HexColor('#001E57'),spaceBefore=14));story=[]
    story += [Paragraph('DUARTE REPORTS',styles['DRTitle']),Paragraph('Relatório operacional • E-mails e matrizes',styles['Heading2']),Spacer(1,7*mm)]
    for label,val in [('Período analisado',meta.get('period','Não informado')),('Gerado em',datetime.now().astimezone().strftime('%d/%m/%Y %H:%M %Z')),('Responsável',meta.get('responsible') or 'Não informado'),('Fonte efetiva',meta.get('source') or 'Importação local')]:story.append(Paragraph(f'<b>{label}:</b> {val}',styles['BodyText']))
    story += [Paragraph('Resumo executivo',styles['DRH']),Paragraph(f"Foram analisadas <b>{len(emails)}</b> mensagens únicas. Há <b>{len({e.get('cliente') for e in emails if e.get('cliente')})}</b> clientes/operadoras explicitamente identificados na fonte. Duplicidades removidas: <b>{payload.get('duplicates',0)}</b>.",styles['BodyText'])]
    story += [Paragraph('Detalhamento por cliente/operadora',styles['DRH'])]
    grouped={}
    for e in emails:
        k=e.get('cliente') or 'Não identificado';g=grouped.setdefault(k,[0,'']);g[0]+=1;g[1]=max(g[1],e.get('data') or '')
    table=[['Cliente/operadora','Mensagens','Última data']]+[[k,str(v[0]),v[1] or 'Não identificada'] for k,v in sorted(grouped.items())]
    t=Table(table,repeatRows=1,colWidths=[85*mm,28*mm,55*mm]);t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#001E57')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,-1),'Helvetica'),('FONTSIZE',(0,0),(-1,-1),8),('GRID',(0,0),(-1,-1),.3,colors.HexColor('#D7E0EB')),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F3F6FA')]),('VALIGN',(0,0),(-1,-1),'TOP')]))
    story.append(t);story += [Paragraph('Pontos de atenção',styles['DRH']),Paragraph('Fatos calculados: contagens são derivadas exclusivamente dos metadados importados e do período aplicado. Associação por cliente só é considerada quando veio informada na fonte ou foi revisada; destinatários e assunto, isoladamente, não confirmam atendimento.',styles['BodyText']),Paragraph(f"Matriz importada: {matrix.get('arquivo','Não importada')}. Registros distintos identificados: {matrix.get('registros_unicos',0)}; códigos únicos: {matrix.get('codigos_unicos',0)}. Estes dados não foram comparados com Duarte Performance.",styles['BodyText']),Paragraph('Metodologia e limitações',styles['DRH']),Paragraph('Deduplicação por Message-ID quando disponível; caso contrário, hash de data, remetente, destinatários e assunto. Respostas continuam sendo mensagens distintas; o identificador de conversa é exibido, mas não somado como mensagem. Sem conexão Outlook/Graph, sem autenticação e sem persistência central neste pacote. Nenhuma avaliação automática de produtividade é produzida.',styles['BodyText'])]
    doc.build(story);return buf.getvalue()

def parse_multipart_form(content_type, body):
    """Parse multipart/form-data using only the Python standard library."""
    message=BytesParser(policy=policy.default).parsebytes(
        f'Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n'.encode('ascii')+body
    )
    if not message.is_multipart():
        raise ValueError('Envie os arquivos como multipart/form-data.')
    fields={}; files={}
    for part in message.iter_parts():
        if part.get_content_disposition()!='form-data':
            continue
        name=part.get_param('name',header='content-disposition')
        if not name:
            continue
        filename=part.get_filename()
        value=part.get_payload(decode=True) or b''
        if filename is not None:
            files.setdefault(name,[]).append({'filename':filename,'data':value})
        else:
            charset=part.get_content_charset() or 'utf-8'
            fields[name]=value.decode(charset,errors='replace')
    return fields,files

class Handler(SimpleHTTPRequestHandler):
    def __init__(self,*a,**kw):super().__init__(*a,directory=str(ROOT),**kw)
    def log_message(self,*args):pass
    def send_json(self,obj,status=200):
        b=json_bytes(obj);self.send_response(status);self.send_header('Content-Type','application/json; charset=utf-8');self.send_header('Content-Length',str(len(b)));self.end_headers();self.wfile.write(b)
    def do_GET(self):
        if self.path=='/api/health':return self.send_json({'ok':True,'mode':'local-file-import','auth':'not-configured','graph':'not-configured','calls':'future'})
        return super().do_GET()
    def do_POST(self):
        if not self.path.startswith('/api/'):return self.send_error(404)
        if self.path not in ('/api/emails','/api/matrix','/api/export/xlsx','/api/export/pdf'):
            return self.send_json({'error':'Operação não encontrada'},404)
        try:
            length=int(self.headers.get('Content-Length','0'))
            if length>MAX_UPLOAD+1024*1024:raise ValueError('A solicitação excede o limite de 25 MB.')
            body=self.rfile.read(length)
            if self.path in ('/api/export/xlsx','/api/export/pdf'):
                payload=json.loads(body or b'{}')
                content=make_xlsx(payload) if self.path.endswith('xlsx') else make_pdf(payload)
                fname='DuarteReports_Relatorio.xlsx' if self.path.endswith('xlsx') else 'DuarteReports_Relatorio.pdf'
                self.send_response(200);self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if fname.endswith('xlsx') else 'application/pdf');self.send_header('Content-Disposition',f'attachment; filename="{fname}"');self.send_header('Content-Length',str(len(content)));self.end_headers();self.wfile.write(content);return
            fields,files=parse_multipart_form(self.headers.get('Content-Type',''),body)
            if self.path=='/api/emails':
                items=files.get('files',[])
                if not items:raise ValueError('Selecione ao menos um arquivo para importar.')
                allrows=[];names=[]
                for item in items:
                    name=os.path.basename(item['filename'] or 'importacao');data=item['data']
                    if len(data)>MAX_UPLOAD:raise ValueError('Arquivo excede o limite de 25 MB.')
                    names.append(name)
                    if Path(name).suffix.lower()=='.eml':allrows.append(parse_eml(name,data))
                    else:allrows.extend(parse_rows(name,data))
                rows,dups=email_normalize(allrows,', '.join(names))
                start=fields.get('start','');end=fields.get('end','')
                def within(x):
                    d=(x.get('data') or '')[:10]
                    return (not start or d>=start) and (not end or d<=end)
                rows=[x for x in rows if within(x)]
                return self.send_json({'emails':rows,'duplicates':dups,'files':names,'source':'Importação local: '+', '.join(names),'count':len(rows)})
            items=files.get('file',[])
            if not items:raise ValueError('Selecione o arquivo da matriz.')
            item=items[0];name=os.path.basename(item['filename'] or 'importacao');data=item['data']
            if len(data)>MAX_UPLOAD:raise ValueError('Arquivo excede o limite de 25 MB.')
            return self.send_json(matrix_parse(name,data))
        except Exception as e:return self.send_json({'error':str(e) or 'Falha ao processar arquivo'},400)

if __name__=='__main__':
    port=int(os.environ.get('PORT','8000'));print(f'Duarte Reports disponível na porta {port}',flush=True);ThreadingHTTPServer(('0.0.0.0',port),Handler).serve_forever()
