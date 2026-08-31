import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title='库存差异分析', layout='wide')
st.title('WMS 与 EOP 库存差异分析')

uploaded = st.file_uploader('上传包含【库存明细】与【EOP】工作表的 Excel 文件', type=['xlsx'])

def analyze(file):
    wms = pd.read_excel(file, sheet_name='库存明细')
    eop = pd.read_excel(file, sheet_name='EOP')
    
    # 1. Clean whitespace from strings and column names
    wms.columns = wms.columns.astype(str).str.strip()
    eop.columns = eop.columns.astype(str).str.strip()
    
    wms['SKU'] = wms['SKU'].astype(str).str.strip()
    wms['条码1'] = wms['条码1'].astype(str).str.strip()
    wms['商品名称'] = wms['商品名称'].astype(str).str.strip()
    
    eop['商品编码'] = eop['商品编码'].astype(str).str.strip()
    eop['商品条码'] = eop['商品条码'].astype(str).str.strip()
    eop['商品名称'] = eop['商品名称'].astype(str).str.strip()
    
    # 2. Extract a single primary name per SKU to eliminate splitting
    all_names = pd.concat([
        wms[['SKU', '商品名称']].rename(columns={'SKU': '商品编码'}),
        eop[['商品编码', '商品名称']]
    ])
    name_map = all_names.groupby('商品编码')['商品名称'].first().to_dict()
    
    # 3. Group strictly by product code to prevent duplicate row splits
    wms_grp = wms.groupby('SKU', dropna=False, as_index=False).agg({
        '库存数量': 'sum',
        '条码1': 'first',
        '客户ID': lambda x: '+'.join(sorted(set(x.astype(str))))
    }).rename(columns={'SKU': '商品编码', '条码1': '商品条码', '客户ID': 'WMS客户', '库存数量': 'WMS库存数量'})
    
    eop_grp = eop.groupby('商品编码', dropna=False, as_index=False).agg({
        '期末数量': 'sum',
        '商品条码': 'first',
        '大类': 'first',
        '商品新分类': 'first'
    }).rename(columns={'期末数量': 'EOP期末数量'})
    
    # 4. Merge results strictly on SKU (商品编码)
    result = pd.merge(eop_grp, wms_grp, on='商品编码', how='outer', suffixes=('', '_WMS'))
    
    # Fill in standardized names and missing attributes
    result['商品条码'] = result['商品条码'].fillna(result.get('商品条码_WMS')).fillna('')
    result['商品名称'] = result['商品编码'].map(name_map).fillna('')
    result['大类'] = result['大类'].fillna('unknown')
    result['商品新分类'] = result['商品新分类'].fillna('unknown')
    result['WMS客户'] = result['WMS客户'].fillna('')
    result['EOP期末数量'] = result['EOP期末数量'].fillna(0)
    result['WMS库存数量'] = result['WMS库存数量'].fillna(0)
    
    result['匹配方式'] = '商品编码'
    result['差异 (WMS-EOP)'] = result['WMS库存数量'] - result['EOP期末数量']
    result['绝对差异'] = result['差异 (WMS-EOP)'].abs()

    def status(r):
        if r['EOP期末数量'] == 0 and r['WMS库存数量'] > 0:
            return '仅WMS'
        if r['WMS库存数量'] == 0 and r['EOP期末数量'] > 0:
            return '仅EOP'
        if r['差异 (WMS-EOP)'] > 0:
            return 'WMS超出'
        if r['差异 (WMS-EOP)'] < 0:
            return 'WMS短少'
        return '一致'

    result['状态'] = result.apply(status, axis=1)
    result = result.sort_values('绝对差异', ascending=False)
    
    cols = ['商品编码', '商品条码', '商品名称', '大类', '商品新分类', 'WMS客户', '匹配方式', 'EOP期末数量', 'WMS库存数量', '差异 (WMS-EOP)', '绝对差异', '状态']
    return result[cols]

if uploaded:
    result = analyze(uploaded)
    st.metric('差异SKU数', int((result['差异 (WMS-EOP)'] != 0).sum()))
    st.dataframe(result, use_container_width=True)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result.to_excel(writer, sheet_name='库存差异分析', index=False)
    
    st.download_button('下载分析结果Excel', output.getvalue(), file_name='库存差异分析结果.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
