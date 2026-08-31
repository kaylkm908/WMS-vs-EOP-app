import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title='库存差异分析', layout='wide')
st.title('WMS 与 EOP 库存差异分析')

uploaded = st.file_uploader('上传包含【库存明细】与【EOP】工作表的 Excel 文件', type=['xlsx'])


def analyze(file):
    wms = pd.read_excel(file, sheet_name='库存明细')
    eop = pd.read_excel(file, sheet_name='EOP')

    wms_grp = wms.groupby(['SKU','条码1','商品名称'], dropna=False, as_index=False).agg({
        '库存数量':'sum',
        '客户ID': lambda x: '+'.join(sorted(set(x.astype(str))))
    })
    wms_grp = wms_grp.rename(columns={'SKU':'商品编码','条码1':'商品条码','客户ID':'WMS客户','库存数量':'WMS库存数量'})

    eop_grp = eop.groupby(['商品编码','商品条码','商品名称','大类','商品新分类'], dropna=False, as_index=False)['期末数量'].sum()
    eop_grp = eop_grp.rename(columns={'期末数量':'EOP期末数量'})

    result = pd.merge(eop_grp, wms_grp, on='商品编码', how='outer', suffixes=('','_WMS'))

    result['商品条码'] = result['商品条码'].fillna(result.get('商品条码_WMS'))
    result['商品名称'] = result['商品名称'].fillna(result.get('商品名称_WMS'))
    result['大类'] = result['大类'].fillna('unknown')
    result['商品新分类'] = result['商品新分类'].fillna('unknown')
    result['WMS客户'] = result['WMS客户'].fillna('')
    result['EOP期末数量'] = result['EOP期末数量'].fillna(0)
    result['WMS库存数量'] = result['WMS库存数量'].fillna(0)
    result['匹配方式'] = result['商品编码'].apply(lambda x: '商品编码')

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

    cols = ['商品编码','商品条码','商品名称','大类','商品新分类','WMS客户','匹配方式','EOP期末数量','WMS库存数量','差异 (WMS-EOP)','绝对差异','状态']
    return result[cols]

if uploaded:
    result = analyze(uploaded)
    st.metric('差异SKU数', int((result['差异 (WMS-EOP)']!=0).sum()))
    st.dataframe(result, use_container_width=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result.to_excel(writer, sheet_name='库存差异分析', index=False)

    st.download_button('下载分析结果Excel', output.getvalue(), file_name='库存差异分析结果.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
