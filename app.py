import sqlite3
from datetime import date
from io import BytesIO
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
from reportlab.lib.utils import ImageReader


APP_DIR=Path(__file__).resolve().parent; DB=APP_DIR/'stock_tracker.db'
st.set_page_config(page_title='Stock Profit Tracker Pro',page_icon='📈',layout='wide')

def conn():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c

def init():
    c=conn(); c.execute('''CREATE TABLE IF NOT EXISTS transactions(id INTEGER PRIMARY KEY AUTOINCREMENT,person TEXT NOT NULL,stock_name TEXT NOT NULL,sell_date TEXT NOT NULL,quantity INTEGER NOT NULL CHECK(quantity>0),buying_price REAL NOT NULL CHECK(buying_price>=0),selling_price REAL NOT NULL CHECK(selling_price>=0),created_at TEXT DEFAULT CURRENT_TIMESTAMP)''')
    if c.execute('SELECT COUNT(*) FROM transactions').fetchone()[0]==0:
        c.executemany('INSERT INTO transactions(person,stock_name,sell_date,quantity,buying_price,selling_price) VALUES(?,?,?,?,?,?)',[
            ('Papa','HFCL','2026-05-13',1000,17.4,147.4),('Papa','Indo MIM','2026-07-30',120,485,707),('Papa','Indo MIM','2026-07-30',100,485,719),('Papa','Indo MIM','2026-09-22',100,485,1173.4),('Papa','Glasswall Systems','2026-09-22',648,182,307.5)])
    c.commit(); c.close()

def getdf():
    c=conn(); rows=c.execute('SELECT * FROM transactions ORDER BY sell_date DESC,id DESC').fetchall(); c.close()
    a=[]
    for r in rows:
        p=round(r['selling_price']-r['buying_price'],2); t=round(p*r['quantity'],2)
        a.append({'ID':r['id'],'Name':r['person'],'Stock Name':r['stock_name'],'Sell Date':pd.to_datetime(r['sell_date']).date(),'Quantity':r['quantity'],'Buying Price':r['buying_price'],'Selling Price':r['selling_price'],'Profit per Share':p,'Total Profit':t})
    return pd.DataFrame(a)

def add(person,stock,d,q,b,s):
    c=conn(); c.execute('INSERT INTO transactions(person,stock_name,sell_date,quantity,buying_price,selling_price) VALUES(?,?,?,?,?,?)',(person,stock.strip(),d.isoformat(),int(q),float(b),float(s))); c.commit(); c.close()

def delete(i):
    c=conn(); c.execute('DELETE FROM transactions WHERE id=?',(int(i),)); c.commit(); c.close()

def clear():
    c=conn(); c.execute('DELETE FROM transactions'); c.commit(); c.close()

def money(x): return f'₹{x:,.2f}'

def excel(df):
    out=BytesIO(); x=df.drop(columns=['ID']).copy(); x['Sell Date']=pd.to_datetime(x['Sell Date']).dt.date
    with pd.ExcelWriter(out,engine='openpyxl') as w:
        x.to_excel(w,index=False,sheet_name='Transactions'); wb=w.book; ws=w.sheets['Transactions']
        fill=PatternFill('solid',fgColor='1F4E78'); font=Font(color='FFFFFF',bold=True); side=Side(style='thin',color='D9E1F2')
        for cell in ws[1]: cell.fill=fill; cell.font=font; cell.alignment=Alignment(horizontal='center'); cell.border=Border(bottom=side)
        for row in ws.iter_rows(min_row=2):
            row[2].number_format='dd mmm yyyy'; row[3].number_format='#,##0'
            for j in [4,5,6,7]: row[j].number_format='₹#,##0.00'
        tr=ws.max_row+2; ws.cell(tr,1,'TOTAL'); ws.cell(tr,8,f'=SUM(H2:H{tr-2})'); ws.cell(tr,8).number_format='₹#,##0.00'
        for c in ws[tr]: c.font=Font(bold=True); c.fill=PatternFill('solid',fgColor='E2F0D9')
        for i,wid in enumerate([14,24,16,12,17,17,19,17],1): ws.column_dimensions[get_column_letter(i)].width=wid
        ws.freeze_panes='A2'; ws.auto_filter.ref=f'A1:H{tr-2}'
        s=wb.create_sheet('Summary'); inv=float((x['Buying Price']*x['Quantity']).sum()); sales=float((x['Selling Price']*x['Quantity']).sum()); profit=float(x['Total Profit'].sum())
        vals=[['Stock Profit Tracker — Summary',''],['Transactions',len(x)],['Total Quantity',int(x['Quantity'].sum())],['Total Investment',inv],['Total Sales Value',sales],['Total Profit',profit]]
        for r in vals:s.append(r)
        for c in s[1]:c.fill=fill;c.font=Font(color='FFFFFF',bold=True,size=14)
        for r in range(2,s.max_row+1):s.cell(r,2).number_format='₹#,##0.00'
        s.column_dimensions['A'].width=26;s.column_dimensions['B'].width=22
    return out.getvalue()

def pdf(df):
    """Create a PDF containing KPI summary, analytics tables/charts, and transactions."""
    import tempfile
    import os
    import matplotlib.pyplot as plt

    out=BytesIO()
    doc=SimpleDocTemplate(
        out,
        pagesize=landscape(A4),
        rightMargin=8*mm,leftMargin=8*mm,topMargin=8*mm,bottomMargin=8*mm
    )
    stl=getSampleStyleSheet()
    title=ParagraphStyle('t',parent=stl['Title'],alignment=TA_CENTER,fontSize=18)
    h2=ParagraphStyle('h2',parent=stl['Heading2'],fontSize=13,spaceBefore=5*mm,spaceAfter=2*mm)
    small=ParagraphStyle('small',parent=stl['BodyText'],fontSize=8)

    inv=float((df['Buying Price']*df['Quantity']).sum()) if len(df) else 0
    sales=float((df['Selling Price']*df['Quantity']).sum()) if len(df) else 0
    profit=float(df['Total Profit'].sum()) if len(df) else 0
    qty=int(df['Quantity'].sum()) if len(df) else 0
    ret=(profit/inv*100) if inv else 0
    profitable=int((df['Total Profit']>0).sum())
    loss=int((df['Total Profit']<0).sum())
    breakeven=int((df['Total Profit']==0).sum())

    story=[
        Paragraph('Stock Profit Tracker — Profit & Analytics Report',title),
        Spacer(1,4*mm)
    ]

    # KPI summary
    sm=Table([
        ['Transactions','Quantity','Investment','Sales Value','Total Profit','Return'],
        [str(len(df)),f"{qty:,}",money(inv),money(sales),money(profit),f"{ret:.2f}%"]
    ],colWidths=[30*mm,30*mm,45*mm,45*mm,45*mm,30*mm])
    sm.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1F4E78')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('BACKGROUND',(0,1),(-1,1),colors.HexColor('#E2F0D9')),
        ('GRID',(0,0),(-1,-1),.4,colors.grey),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('FONTSIZE',(0,0),(-1,-1),9)
    ]))
    story += [sm, Spacer(1,3*mm)]

    status=Table([
        ['Profitable Transactions','Loss Transactions','Break-even Transactions'],
        [str(profitable),str(loss),str(breakeven)]
    ],colWidths=[55*mm,55*mm,55*mm])
    status.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#D9EAF7')),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('GRID',(0,0),(-1,-1),.4,colors.grey),
        ('ALIGN',(0,0),(-1,-1),'CENTER')
    ]))
    story += [status]

    # Analytics tables
    person=df.groupby('Name',as_index=False).agg(
        Transactions=('ID','count'),
        Quantity=('Quantity','sum'),
        Investment=('Buying Price',lambda x: 0)
    )
    # Recalculate investment/sales at row level so grouped values are accurate.
    tmp=df.copy()
    tmp['Investment']=tmp['Buying Price']*tmp['Quantity']
    tmp['Sales Value']=tmp['Selling Price']*tmp['Quantity']
    person=tmp.groupby('Name',as_index=False).agg(
        Transactions=('ID','count'),
        Quantity=('Quantity','sum'),
        Investment=('Investment','sum'),
        Sales=('Sales Value','sum'),
        Profit=('Total Profit','sum')
    )

    stock=tmp.groupby('Stock Name',as_index=False).agg(
        Transactions=('ID','count'),
        Quantity=('Quantity','sum'),
        Investment=('Investment','sum'),
        Sales=('Sales Value','sum'),
        Profit=('Total Profit','sum')
    )

    tmp['Month']=pd.to_datetime(tmp['Sell Date']).dt.to_period('M').astype(str)
    monthly=tmp.groupby('Month',as_index=False).agg(
        Transactions=('ID','count'),
        Quantity=('Quantity','sum'),
        Investment=('Investment','sum'),
        Sales=('Sales Value','sum'),
        Profit=('Total Profit','sum')
    )

    def analytics_table(frame, first_col, widths):
        headers=[first_col,'Transactions','Quantity','Investment','Sales','Profit']
        rows=[headers]
        for _,r in frame.iterrows():
            rows.append([
                str(r[first_col]),
                f"{int(r['Transactions']):,}",
                f"{int(r['Quantity']):,}",
                money(r['Investment']),
                money(r['Sales']),
                money(r['Profit'])
            ])
        tb=Table(rows,colWidths=widths,repeatRows=1)
        tb.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1F4E78')),
            ('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
            ('GRID',(0,0),(-1,-1),.3,colors.grey),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F5F8FA')]),
            ('FONTSIZE',(0,0),(-1,-1),8),
            ('ALIGN',(1,1),(-1,-1),'RIGHT')
        ]))
        return tb

    story += [
        Paragraph('Analytics — Person-wise',h2),
        analytics_table(person,'Name',[40*mm,30*mm,30*mm,45*mm,45*mm,45*mm]),
        Paragraph('Analytics — Stock-wise',h2),
        analytics_table(stock,'Stock Name',[40*mm,30*mm,30*mm,45*mm,45*mm,45*mm]),
        Paragraph('Analytics — Monthly',h2),
        analytics_table(monthly,'Month',[40*mm,30*mm,30*mm,45*mm,45*mm,45*mm])
    ]

    # Generate analytics charts temporarily and embed them in the PDF.
    with tempfile.TemporaryDirectory() as td:
        chart_paths=[]

        def save_bar(labels, values, title_text, filename, ylabel='Profit (₹)'):
            if len(labels)==0:
                return None
            fig,ax=plt.subplots(figsize=(8.5,3.2))
            ax.bar(labels, values)
            ax.set_title(title_text)
            ax.set_ylabel(ylabel)
            ax.tick_params(axis='x',rotation=35)
            fig.tight_layout()
            path=os.path.join(td,filename)
            fig.savefig(path,dpi=150,bbox_inches='tight')
            plt.close(fig)
            return path

        def save_line(labels, values, title_text, filename):
            if len(labels)==0:
                return None
            fig,ax=plt.subplots(figsize=(8.5,3.2))
            ax.plot(labels, values, marker='o')
            ax.set_title(title_text)
            ax.set_ylabel('Profit (₹)')
            ax.tick_params(axis='x',rotation=35)
            fig.tight_layout()
            path=os.path.join(td,filename)
            fig.savefig(path,dpi=150,bbox_inches='tight')
            plt.close(fig)
            return path

        p=save_bar(person['Name'].tolist(),person['Profit'].tolist(),'Profit by Person','profit_by_person.png')
        if p: chart_paths.append(p)
        p=save_bar(stock['Stock Name'].tolist(),stock['Profit'].tolist(),'Profit by Stock','profit_by_stock.png')
        if p: chart_paths.append(p)
        p=save_line(monthly['Month'].tolist(),monthly['Profit'].tolist(),'Monthly Profit Trend','monthly_profit.png')
        if p: chart_paths.append(p)

        story.append(PageBreak())
        story.append(Paragraph('Analytics Charts',h2))
        for cp in chart_paths:
            story.append(RLImage(cp,width=250*mm,height=90*mm))
            story.append(Spacer(1,3*mm))

        story.append(PageBreak())
        story.append(Paragraph('Detailed Transactions',h2))
        rows=[['Name','Stock','Sell Date','Qty','Buy','Sell','Profit/Share','Total Profit']]
        for _,r in df.iterrows():
            rows.append([
                r['Name'],r['Stock Name'],
                pd.to_datetime(r['Sell Date']).strftime('%d %b %Y'),
                f"{int(r['Quantity']):,}",
                money(r['Buying Price']),money(r['Selling Price']),
                money(r['Profit per Share']),money(r['Total Profit'])
            ])
        tb=Table(rows,repeatRows=1,colWidths=[23*mm,42*mm,29*mm,18*mm,28*mm,28*mm,35*mm,38*mm])
        tb.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1F4E78')),
            ('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
            ('GRID',(0,0),(-1,-1),.3,colors.grey),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F5F8FA')]),
            ('FONTSIZE',(0,0),(-1,-1),8),
            ('ALIGN',(2,1),(-1,-1),'RIGHT')
        ]))
        story += [
            tb,
            Spacer(1,4*mm),
            Paragraph(
                'Profit per Share = Selling Price − Buying Price. '
                'Total Profit = Profit per Share × Quantity. '
                'No brokerage, STT, GST, taxes, or other charges are included.',
                small
            )
        ]
        doc.build(story)

    return out.getvalue()


init(); df=getdf(); st.title('📈 Stock Profit Tracker Pro'); st.caption('Local-first stock transaction tracker with automatic calculations, audit checks, Excel and PDF reports.')
with st.sidebar:
    st.header('➕ Add Transaction')
    with st.form('add',clear_on_submit=True):
        person=st.text_input('Name',placeholder='Enter any name — Papa, Mummy, Naman, Rahul, etc.'); stock=st.text_input('Stock Name'); d=st.date_input('Sell Date',date.today()); q=st.number_input('Quantity',1,10000000,1,1); b=st.number_input('Buying Price (₹)',0.0,100000000.0,0.0,.10); s=st.number_input('Selling Price (₹)',0.0,100000000.0,0.0,.10); ok=st.form_submit_button('Add Transaction',type='primary',use_container_width=True)
        if ok:
            if not stock.strip(): st.error('Stock name is required.')
            else: add(person,stock,d,q,b,s); st.success(f'Added. Profit/share: {money(s-b)} | Total: {money((s-b)*q)}'); st.rerun()
    st.divider(); st.header('📥 Import')
    up=st.file_uploader('Excel/CSV',type=['xlsx','csv'])
    if up and st.button('Import Rows',use_container_width=True):
        try:
            imp=pd.read_csv(up) if up.name.lower().endswith('.csv') else pd.read_excel(up)
            req=['Name','Stock Name','Sell Date','Quantity','Buying Price','Selling Price']; miss=[x for x in req if x not in imp.columns]
            if miss: st.error('Missing: '+', '.join(miss))
            else:
                c=conn()
                for _,r in imp.iterrows(): c.execute('INSERT INTO transactions(person,stock_name,sell_date,quantity,buying_price,selling_price) VALUES(?,?,?,?,?,?)',(str(r['Name']),str(r['Stock Name']),pd.to_datetime(r['Sell Date']).date().isoformat(),int(float(r['Quantity'])),float(r['Buying Price']),float(r['Selling Price'])))
                c.commit();c.close();st.success(f'Imported {len(imp)} rows.');st.rerun()
        except Exception as e: st.error(str(e))
    st.divider(); st.header('🧹 Data Management')
    if st.button('Delete All Transactions',use_container_width=True): st.session_state.confirm=True
    if st.session_state.get('confirm'):
        st.warning('This permanently deletes the local transaction database.'); a,bx=st.columns(2)
        if a.button('Confirm'): clear();st.session_state.confirm=False;st.rerun()
        if bx.button('Cancel'): st.session_state.confirm=False;st.rerun()

profit=float(df['Total Profit'].sum()) if len(df) else 0; inv=float((df['Buying Price']*df['Quantity']).sum()) if len(df) else 0; sales=float((df['Selling Price']*df['Quantity']).sum()) if len(df) else 0; qty=int(df['Quantity'].sum()) if len(df) else 0
k=st.columns(5); k[0].metric('Total Profit',money(profit)); k[1].metric('Investment',money(inv)); k[2].metric('Sales Value',money(sales)); k[3].metric('Quantity',f'{qty:,}'); k[4].metric('Return',f'{profit/inv*100:.2f}%' if inv else '0.00%')
t1,t2,t3=st.tabs(['📋 Transactions','📊 Analytics','📤 Export'])
with t1:
    if len(df):
        v=df.copy();v['Sell Date']=pd.to_datetime(v['Sell Date']).dt.strftime('%d %b %Y'); st.dataframe(v.drop(columns=['ID']),use_container_width=True,hide_index=True)
        opts={f"#{int(r['ID'])} — {r['Stock Name']} — {r['Sell Date']} — {money(r['Total Profit'])}":int(r['ID']) for _,r in df.iterrows()}; sel=st.selectbox('Delete transaction',list(opts))
        if st.button('Delete Selected'): delete(opts[sel]);st.rerun()
    else: st.info('No transactions yet.')
with t2:
    if len(df):
        a,b=st.columns(2); ps=df.groupby('Name')['Total Profit'].sum(); ss=df.groupby('Stock Name')['Total Profit'].sum(); a.subheader('Profit by Person');a.bar_chart(ps);b.subheader('Profit by Stock');b.bar_chart(ss)
        m=df.copy();m['Month']=pd.to_datetime(m['Sell Date']).dt.to_period('M').astype(str);st.subheader('Monthly Profit');st.line_chart(m.groupby('Month')['Total Profit'].sum())
        audit=df.copy();audit['Calculated Profit/Share']=(audit['Selling Price']-audit['Buying Price']).round(2);audit['Calculated Total Profit']=(audit['Calculated Profit/Share']*audit['Quantity']).round(2);audit['Profit/Share OK']=audit['Profit per Share'].round(2)==audit['Calculated Profit/Share'];audit['Total Profit OK']=audit['Total Profit'].round(2)==audit['Calculated Total Profit'];st.subheader('Calculation Audit');st.dataframe(audit[['Stock Name','Profit per Share','Calculated Profit/Share','Profit/Share OK','Total Profit','Calculated Total Profit','Total Profit OK']],use_container_width=True,hide_index=True);st.success('All calculations are correct.') if audit['Profit/Share OK'].all() and audit['Total Profit OK'].all() else st.error('Calculation mismatch detected.')
    else: st.info('Add transactions to see analytics.')
with t3:
    if len(df):
        c1,c2=st.columns(2);c1.download_button('⬇️ Download Excel',excel(df),f'Stock_Profit_Report_{date.today()}.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True);c2.download_button('⬇️ Download PDF',pdf(df),f'Stock_Profit_Report_{date.today()}.pdf','application/pdf',use_container_width=True);st.download_button('Download CSV Backup',df.drop(columns=['ID']).to_csv(index=False).encode(),f'Stock_Profit_Backup_{date.today()}.csv','text/csv',use_container_width=True)
    else: st.info('Add transactions before exporting.')
st.divider();st.caption('Data is stored locally in stock_tracker.db. Profit = Selling Price − Buying Price; Total Profit = Profit/Share × Quantity. No brokerage, STT, GST, taxes, or other charges are included.')
