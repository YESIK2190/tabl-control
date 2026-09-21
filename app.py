import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Dashboard Comparativo", layout="wide")
st.title("Dashboard Comparativo de Operaciones")

def normalizar(df):
    df.columns=(df.columns.astype(str).str.strip().str.upper().str.replace('Ó','O',regex=False))
    return df

base_file=st.sidebar.file_uploader('BASE.xlsx',type=['xlsx'])
fin_file=st.sidebar.file_uploader('DESACRGA FINANSOFT.xlsx',type=['xlsx'])
val_file=st.sidebar.file_uploader('VALIDACION VENTAS.xlsx',type=['xlsx'])

if base_file and fin_file and val_file:
    try:
        base=pd.read_excel(base_file)
        base=normalizar(base)
        if 'CODIGO' not in base.columns:
            base=pd.read_excel(base_file,header=4)
            base=normalizar(base)

        fins=normalizar(pd.read_excel(fin_file))
        valid=normalizar(pd.read_excel(val_file))

        for nombre,df in [('BASE',base),('FINANSOFT',fins),('VALIDACION',valid)]:
            if 'CODIGO' not in df.columns:
                st.error(f'No existe la columna CODIGO en {nombre}')
                st.write(df.columns.tolist())
                st.stop()

        base['CODIGO']=pd.to_numeric(base['CODIGO'],errors='coerce')
        fins['CODIGO']=pd.to_numeric(fins['CODIGO'],errors='coerce')
        valid['CODIGO']=pd.to_numeric(valid['CODIGO'],errors='coerce')

        comunes=set(base['CODIGO'].dropna()) & set(fins['CODIGO'].dropna()) & set(valid['CODIGO'].dropna())
        c1,c2,c3,c4=st.columns(4)
        c1.metric('BASE',len(base))
        c2.metric('FINANSOFT',len(fins))
        c3.metric('VALIDACION',len(valid))
        c4.metric('COINCIDENCIAS',len(comunes))

        cruce=base[['CODIGO']].merge(valid[['CODIGO']],on='CODIGO',how='outer',indicator=True)
        resumen=cruce['_merge'].value_counts().reset_index()
        resumen.columns=['Estado','Cantidad']
        st.plotly_chart(px.bar(resumen,x='Estado',y='Cantidad',color='Estado'),use_container_width=True)

        if 'RESPONSABLE' in base.columns:
            usr=base.groupby('RESPONSABLE').size().reset_index(name='OPERACIONES')
            st.plotly_chart(px.bar(usr,x='RESPONSABLE',y='OPERACIONES'),use_container_width=True)

        if 'FIDEICOMISO' in fins.columns:
            fid=fins.groupby('FIDEICOMISO').size().reset_index(name='OPERACIONES').sort_values('OPERACIONES',ascending=False).head(20)
            st.plotly_chart(px.bar(fid,x='OPERACIONES',y='FIDEICOMISO',orientation='h'),use_container_width=True)

        dif=cruce[cruce['_merge']!='both']
        output=BytesIO()
        with pd.ExcelWriter(output,engine='xlsxwriter') as writer:
            resumen.to_excel(writer,sheet_name='Resumen',index=False)
            dif.to_excel(writer,sheet_name='Diferencias',index=False)

        st.download_button('Descargar Resultado Excel',output.getvalue(),'Resultado_Cruce.xlsx')
    except Exception as e:
        st.exception(e)
else:
    st.info('Cargue los tres archivos Excel')
