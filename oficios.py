import streamlit as st
import pandas as pd

# Configurar la página
st.set_page_config(page_title="Registro de Oficios", page_icon="📄", layout="wide")

# Cargar el archivo con la ruta específica
file_path = "OFICIOS.2.xlsx"
xls = pd.ExcelFile(file_path)

# Estilos CSS
st.markdown("""
    <style>
        /* Estilos generales */
        body {background-color: #f4f4f4;}
        .stApp {background-color: #ffffff; padding: 20px; border-radius: 10px; box-shadow: 2px 2px 10px rgba(0,0,0,0.1);}
        .stTitle {color: #2E3B55; text-align: center;}
        
        /* Estilos para el sidebar */
        section[data-testid="stSidebar"] {
            background-color: #2E3B55 !important;
        }
        
        section[data-testid="stSidebar"] .stMarkdown p {
            color: white !important;
        }
        
        section[data-testid="stSidebar"] h1, 
        section[data-testid="stSidebar"] h2, 
        section[data-testid="stSidebar"] h3,
        section[data-testid="stSidebar"] label {
            color: white !important;
        }
        
        /* Estilo específico para el texto dentro del selectbox */
        section[data-testid="stSidebar"] select option,
        section[data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] span {
            color: black !important;
        }
        
        /* Estilos para la tabla */
        table {width: 100%; border-collapse: collapse;}
        th, td {border: 1px solid #ddd; padding: 8px;}
        th {background-color: #4CAF50; color: white; text-align: center !important; font-weight: bold;}
        td {text-align: left;}
        tr:nth-child(even) {background-color: #f2f2f2;}
    </style>
""", unsafe_allow_html=True)

# Selección de hoja
sheet_name = st.selectbox("Selecciona el año", xls.sheet_names)

# Leer el archivo Excel con tipos de datos específicos
df = pd.read_excel(
    xls, 
    sheet_name=sheet_name,
    dtype={'NÚMERO OFICIO': str},  # Asegurar que el número de oficio se lea como texto
    parse_dates=['FECHA']  # Parsear la columna de fecha
)

# Formatear el número de oficio para mantener los ceros iniciales
df['NÚMERO OFICIO'] = df['NÚMERO OFICIO'].str.zfill(3)

# Formatear la fecha para mostrar solo la fecha sin hora
df['FECHA'] = pd.to_datetime(df['FECHA']).dt.strftime('%d/%m/%Y')

# Lista de enlaces por año
enlaces_2024 = {
    "001": "https://drive.google.com/file/d/15vLTUGq3sEw5excGW_NCLVM7GbsRPKON/view?usp=drive_link",
    "002": "https://drive.google.com/file/d/1fShOw0xTL42BW4OUpj0xPz8BfgcjN-y0/view?usp=drive_link",
    "003": "https://drive.google.com/file/d/1vLiK5DEjdHnfhNU88IbZI4m-c9ILMuV4/view?usp=drive_link",
    "004": "https://drive.google.com/file/d/18nY7l7hEHjzqWaOcbpbFh_Ow8NvkznCY/view?usp=drive_link",
    "005": "https://drive.google.com/file/d/1eLRn2cB98eOPG5SX4Z7nyHbYmRtxnbdV/view?usp=drive_link",
    "006": "https://drive.google.com/file/d/1C8e5VMfth79r3uBBkxSWdNVMTUQuA9RV/view?usp=drive_link",
    "007": "https://drive.google.com/open?id=1---FccPfM5Lx9tKoR_mqlq1GUamZB09g",
    "010": "https://drive.google.com/file/d/1dkv7o2tbcVK6H0E4W2J8rdHah7OtV5mI/view?usp=drive_link",
    "011": "https://drive.google.com/file/d/1rA7drPDbVrG8FvoF8YLkarZKvMicmpiZ/view?usp=drive_link",
    "012": "https://drive.google.com/file/d/1UknBYTN2X_-YiX77JHQlA2tM1KnRNetu/view?usp=drive_link",
    "013": "https://drive.google.com/file/d/1lsHWBejAC-IYn9aO0KP09e3xFVRLwuab/view?usp=drive_link",
    "014": "https://drive.google.com/file/d/1XjLU0qEwmePey9IrPjJ-PgjplNgkRT-K/view?usp=drive_link",
    "015": "https://drive.google.com/file/d/1BRJci0rI1rc8Hh-MAE4oL3xvMuLWPgyv/view?usp=drive_link",
    "016": "https://drive.google.com/file/d/1Mi3Ad3KweeA5Sn7n__BzSZWMhfaVosGA/view?usp=drive_link",
    "017": "https://drive.google.com/open?id=1MfIEqmfHLYRfbiF0Xj2E9F3KEGLwTMgx",
    "018": "https://drive.google.com/file/d/17hM4Dm3J4DtlGWBmaYSJn2XTGd2PjTBb/view?usp=drive_link",
    "019": "https://drive.google.com/file/d/1T2FjwzB7WMTZgZc00sh0v3R4ZF9kDEnz/view?usp=drive_link",
    "021": "https://drive.google.com/file/d/1sbiAPHUQeE0dSowCl7IZ0fafAaUQWkoD/view?usp=drive_link",
    "022": "https://drive.google.com/file/d/1cuXR07fY-FVhPj_Ve1mDgS0TXkdWovmj/view?usp=drive_link",
    "023": "https://drive.google.com/file/d/1RbRei71QpY9lCBoR3ZgLVqYElJ76pGOz/view?usp=drive_link",
    "024": "https://drive.google.com/open?id=1Uv51Xd8Uuj-rDCN7zH7zuHU7W2TakC7f",
    "025": "https://drive.google.com/file/d/1RZo5yW-KFx3dKuAD8EaEMd2Il_FK01B2/view?usp=drive_link",
    "026": "https://drive.google.com/file/d/1gq6ulkw5L9-W5wJl7vSYVjaqiu3ZQkZK/view?usp=drive_link",
    "027": "https://drive.google.com/file/d/1rUY-jMY342WAOVuv9x5uEGjKlZSlfnO4/view?usp=drive_link",
    "028": "https://drive.google.com/file/d/1nvpcYPZcnL46mAgRCEXb8krDYvY8hhI0/view?usp=drive_link",
    "030": "https://drive.google.com/open?id=14LX_lnCluJOY91f3biO-NdyF3z16xTRx",
    "031": "https://drive.google.com/file/d/1a77Tsb8HA-X8W3f4W3JoHXFtBjxA2kbF/view?usp=drive_link",
    "032": "https://drive.google.com/file/d/1xeynkvhGVpJn5fSlaVJ_cDCBpZCkhksZ/view?usp=drive_link",
    "033": "https://drive.google.com/file/d/17vPSxKnv8Z4KtdYYg9ySvLSywBxLlfsk/view?usp=drive_link",
    "034": "https://drive.google.com/file/d/1XvbjwTlxsCZyKlVvBmxaVGrjpQIVxZti/view?usp=drive_link",
    "035": "https://drive.google.com/file/d/1qWtxeR0MALyB7757hax5xN3nUlm76HLQ/view?usp=drive_link",
    "036": "https://drive.google.com/file/d/1VNWknfDtHDMoF55HS4kuXXo2fcDtf_eR/view?usp=drive_link",
    "037": "https://drive.google.com/file/d/1beUtrWr6v52DH3cWWB-fVng7gkWPx_IA/view?usp=drive_link",
    "038": "https://drive.google.com/file/d/1fTu_bhAEGze8DFd7flbs2Hrr2dDo5nZe/view?usp=drive_link",
    "041": "https://drive.google.com/file/d/1MJDlJtGaUWAdlGr6p8b9PAfGteUEdPuu/view?usp=drive_link",
    "043": "https://drive.google.com/file/d/1Mi6BB117NbHwZAGJyHjsC4Ep111fCvYq/view?usp=drive_link",
    "044": "https://drive.google.com/file/d/1qqPA-Malodh992jfKPa1kELJQRp_TE4d/view?usp=drive_link",
    "045": "https://drive.google.com/file/d/1kZcY9ohyktAC5U5Sy9pooMaAAb0wTabV/view?usp=drive_link",
    "046": "https://drive.google.com/file/d/19qoBqSwuHJCJZk-gtqC1smXA-0CXOkNX/view?usp=drive_link",
    "047": "https://drive.google.com/file/d/1Z6n_GsAnd_cTbuZ_eRZQTg8b-LT11YxR/view?usp=drive_link",
    "048": "https://drive.google.com/file/d/1T25c7yZSEmXdpnlm1VBZkzlyMl2rHPR9/view?usp=drive_link",
    "049": "https://drive.google.com/file/d/1OCJQKNilg9CVK4GfSfux_PYDkHtcZSu5/view?usp=drive_link",
    "050": "https://drive.google.com/file/d/1PYuD7jqQqGCifbj7_tSmJb7ubvhNoCXY/view?usp=drive_link",
    "051": "https://drive.google.com/file/d/10JSjxheMFxiGb4IX0A6-cR6Dw0AGUwWE/view?usp=drive_link",
    "052": "https://drive.google.com/file/d/1s8IJmLumr33rrsSkYvn3qDjsbl3dqhGa/view?usp=drive_link",
    "053": "https://drive.google.com/file/d/1Nk00pPNUO-aN760e-xxx3ZcUIZbbV2oE/view?usp=drive_link",
    "054": "https://drive.google.com/file/d/1yeTGUM8iNGUGC8Fx6tX87jbAPI4J0-NP/view?usp=drive_link",
    "055": "https://drive.google.com/file/d/11zG0emjhQ0-5H9G9rGt13pkrAWna3Il_/view?usp=drive_link",
    "056": "https://drive.google.com/file/d/1DkmzqTkv4Oyxq3zjWid58_geUowuITK6/view?usp=drive_link",
    "057": "https://drive.google.com/file/d/1Osx6PBaJ1ipxCrI48gUQj1lPoH-kPsQ9/view?usp=drive_link",
    "058": "https://drive.google.com/file/d/1HGA2Qc1kOBDZRNXKQHjbQ7816Novrqyv/view?usp=drive_link",
    "059": "https://drive.google.com/file/d/1_jRmw3CGSl5LNvS3G_EeaJAV3xmL_vkT/view?usp=drive_link",
    "060": "https://drive.google.com/file/d/1zNw5ZRfIHa8m6Ct41YqdGUwaIKPjtxeP/view?usp=drive_link",
    "061": "https://drive.google.com/file/d/1rKb_UmvJ0SKOEsKZroCJRCpRhdnwU4RM/view?usp=drive_link",
    "062": "https://drive.google.com/file/d/1EfgEKYujFqCkrEQKPrcv1raSQIeb66xT/view?usp=drive_link",
    "063": "https://drive.google.com/file/d/1AY3kruKsfDh8zCShzibD7tn8AA6ZECdC/view?usp=drive_link",
    "064": "https://drive.google.com/file/d/1kGJMejVRoXj4cvCItixHDd2Qg1lj6FI7/view?usp=drive_link",
    "065": "https://drive.google.com/file/d/1dZutLfP046S3nH6HKZtBqcuDbKzEgYB9/view?usp=drive_link",
    "067": "https://drive.google.com/file/d/1ArGYlo5qyyF1kJDtHQ-E1dDX91rYXxg-/view?usp=drive_link",
    "068": "https://drive.google.com/file/d/1pSDXwEs9pZYl28FjpgLHxcrbIPcxJ61W/view?usp=drive_link",
    "069": "https://drive.google.com/file/d/1qWjUfKF-KeZlem8mEDicuDGTY_tHisuO/view?usp=drive_link",
    "070": "https://drive.google.com/file/d/1PeMiYbI-jO6TxBjv8q7n01OOqj-00Hv6/view?usp=drive_link",
    "071": "https://drive.google.com/file/d/1OX4GN-LCkqPyAiV68A34CpRfFMEbSA4s/view?usp=drive_link",
    "073": "https://drive.google.com/file/d/1S834G2HC868Hd2dUFV93eRELcKxHrgGV/view?usp=drive_link",
    "074": "https://drive.google.com/file/d/1Azi_z4N_IqBEi1UQAhqoeht9WD8KwdKd/view?usp=drive_link",
    "075": "https://drive.google.com/file/d/16YVuZp9977r4y2sN86ZSd1YvV76oOBum/view?usp=drive_link",
    "076": "https://drive.google.com/file/d/1drdvdiPPVlgPnvKCp92Of7vharMzXmCD/view?usp=drive_link",
    "077": "https://drive.google.com/file/d/162_n47jdulHhqXOh7jnIv6tEKCRtjFR0/view?usp=drive_link",
    "078": "https://drive.google.com/file/d/1fP9qeyRhGtPSPTRPJOJFSWkz4HKXOdNy/view?usp=drive_link",
    "080": "https://drive.google.com/file/d/1DhL4Aw243VL-v4YwHimzYrsKvI-B8xAo/view?usp=drive_link",
    "082": "https://drive.google.com/file/d/1q24-JMsozuXvUyNmHPjnpz80IpLGS1qp/view?usp=drive_link",
    "083": "https://drive.google.com/file/d/1KshJvXzI8iuSl4hGx-BfVv3POkPhYpgu/view?usp=drive_link",
    "084": "https://drive.google.com/file/d/1EzrgPVIMVwhNBsV6S8V5b8K430Aew6FS/view?usp=drive_link",
    "085": "https://drive.google.com/file/d/1m882arPv4d05Xmt8Jl_ojFVsJI03RCVy/view?usp=drive_link",
    "086": "https://drive.google.com/file/d/1_xbNcn2JaC8ofVUfRgZK47JiwIu0MN1s/view?usp=drive_link",
    "087": "https://drive.google.com/file/d/1MiPAD3mBzRr0Nl2qjmfv2YY37clVE6rj/view?usp=drive_link",
    "088": "https://drive.google.com/file/d/1i3VBQwSKefhnJchiT5bYQ_si_i8gUYaB/view?usp=drive_link",
    "089": "https://drive.google.com/file/d/1On2dxt7cl43L705FkCfiHNbsR-jAkS4p/view?usp=drive_link",
    "090": "https://drive.google.com/file/d/1VvQBUMLtVn9BcCTz7xd8Qs1nu7RAxATD/view?usp=drive_link",
    "091": "https://drive.google.com/file/d/1uEYckvSAlugxphAwPxhdWZuBo_z2rO0P/view?usp=drive_link",
    "092": "https://drive.google.com/file/d/1E5VOGhQFyYyDcKjoKZwaNf1pkAQKMdb6/view?usp=drive_link",
    "094": "https://drive.google.com/file/d/1R-mpuCgx53_XSB5P0wvikWFtZXQUqA2J/view?usp=drive_link",
    "095": "https://drive.google.com/file/d/1WMOYn74_RdO_T_e2rxMkzRbZl_2NwHhO/view?usp=drive_link",
    "096": "https://drive.google.com/file/d/1RgOTuoUsg5oN5uBXLi9ovqsX8dj_HDp7/view?usp=drive_link",
    "098": "https://drive.google.com/file/d/1pn5qddt3Asu1v_BDwc_OOhkTRuVgv4TN/view?usp=drive_link",
    "099": "https://drive.google.com/file/d/16Nl3OOn9WxaiqaSV6RRJNGq7aOh6xKA3/view?usp=drive_link",
    "100": "https://drive.google.com/file/d/1JuTdnpsyn9gLO_Fx1fI1TJclJ23ap7Dj/view?usp=drive_link",
    "101": "https://drive.google.com/file/d/1PKj5x1kFj6QGX6Q9PGGoiHQP_wAqh4iC/view?usp=drive_link",
    "102": "https://drive.google.com/file/d/19VebQ2UHNJgwgmRAJHyIHlMxUdF1hGWX/view?usp=drive_link",
    "104": "https://drive.google.com/file/d/1iITh7khtGVml8EcEwOhh3glNdUNw6W1-/view?usp=drive_link",
    "105": "https://drive.google.com/file/d/1e86YRmq4veGXJBN2OfOuAIZk_O4dFnWP/view?usp=drive_link",
    "106": "https://drive.google.com/file/d/1pna2oDHOZVHMk4MjgEWHQ6pDptojiGco/view?usp=drive_link",
    "107": "https://drive.google.com/file/d/1MMmgtpDnFBjPAKSRjHqbjT8IYMYvQfdX/view?usp=drive_link",
    "108": "https://drive.google.com/file/d/1F7enMUkaeS_9ylfbJe3iMhMj96QU06G2/view?usp=drive_link",
    "110": "https://drive.google.com/file/d/1HBcqZvjf8-cAzN09E63nEY6mSVNeGZ4C/view?usp=drive_link",
    "111": "https://drive.google.com/file/d/1cWMZVQ5-hGbslNVR6FEW8wBlr_kw-Fto/view?usp=drive_link",
    "112": "https://drive.google.com/file/d/17OOdbqUOefbygqWgmZnP_VELdPo_KY6m/view?usp=drive_link",
    "113": "https://drive.google.com/file/d/10WZiqpS2wJ_4H3V325x1hkvXuQgQV7E9/view?usp=drive_link",
    "114": "https://drive.google.com/file/d/11gUkSOxuZJ2c7CVqMS-WgbfXMSQcHQnp/view?usp=drive_link",
    "115": "https://drive.google.com/file/d/16V2tRVkkvqpw66L_3fmRE_4THseOtKSI/view?usp=drive_link",
    "116": "https://drive.google.com/file/d/1j1unfzFuSj_Qou3IVY1836phNxn3L2z9/view?usp=drive_link",
    "117": "https://drive.google.com/file/d/1OWrUOpIYE-bLdL5U4CnjGSf5BoxuL6Nn/view?usp=drive_link",
    "118": "https://drive.google.com/file/d/1fhEITnuKZslCGzHKFmPRV_DdlE4UNISN/view?usp=drive_link",
    "119": "https://drive.google.com/file/d/1GDGjJxZTNUE78m3JTbKDylNrk8AAnEMk/view?usp=drive_link",
    "120": "https://drive.google.com/file/d/1t_laQpI1jwvFKL_DoTLTjNhlb25diDSI/view?usp=drive_link",
    "121": "https://drive.google.com/file/d/1iAMqBlkervPjoDugvvHjsxfQuHJ0dzoK/view?usp=drive_link",
    "122": "https://drive.google.com/file/d/1DZyV95EbCpEpmckgyTYIKGYPqkWgxm3K/view?usp=drive_link",
    "123": "https://drive.google.com/file/d/18vrV_DSgdS2DRTrRuspz-SZ-O_p4uSgn/view?usp=drive_link",
    "124": "https://drive.google.com/file/d/1kvTLxV6fAv_rDG_w_oDvAgzBIn8u1nH5/view?usp=drive_link",
    "127": "https://drive.google.com/file/d/1cf6O_2oGibDYAoHhOtgtMesJ4HUJgmKz/view?usp=drive_link",
    "128": "https://drive.google.com/file/d/1MmOw1F0kokdz4slKw3bWCGtk9MuuNwjK/view?usp=drive_link",
    "129": "https://drive.google.com/file/d/1xLSjJACgoR2NDhKn0dTeT_O4ta6Q3zk8/view?usp=drive_link",
    "130": "https://drive.google.com/file/d/1B0frtOfG4NltqFB_S19u1z_uFAMOP4N0/view?usp=drive_link",
    "131": "https://drive.google.com/file/d/14rnXE64pI1YAZLBoN7rgLewdpMdzIICC/view?usp=drive_link",
    "132": "https://drive.google.com/file/d/1ZqbjfM62Qyb0DB5jLQAB33RSEMHiuG-5/view?usp=drive_link",
    "133": "https://drive.google.com/file/d/17yJOcPWvDGaxTmY-PpThKJK9NzKdQ-pp/view?usp=drive_link",
    "134": "https://drive.google.com/file/d/19nHYTz4bKouCi35xRAEO1FU7HlLyJ5ic/view?usp=drive_link",
    "135": "https://drive.google.com/file/d/1Gky4iUG1_buoKAj_8_vCq2_Atlq2Xo9f/view?usp=drive_link",
    "136": "https://drive.google.com/file/d/1BIQnIFZ4cMDCew7MDi5fTpI1Jakf0KOi/view?usp=drive_link",
    "137": "https://drive.google.com/file/d/1XEfc36LDMszWSnb7UDbtIt4j3VhiOeOt/view?usp=drive_link",
    "138": "https://drive.google.com/file/d/1FnKIAYJJt2TpJ3uMIF-FWEWCeESdQMX3/view?usp=drive_link",
    "139": "https://drive.google.com/file/d/11y9F8eHMstRKfxzNYoumjx4dwdVWsJZI/view?usp=drive_link",
    "140": "https://drive.google.com/file/d/1j6TUwMfq2_m7JLqJ0uy0oFd8Aj3AlhBt/view?usp=drive_link",
    "141": "https://drive.google.com/file/d/1Xt434PWSCZoPmTerw6evWAraje3pEfAd/view?usp=drive_link",
    "142": "https://drive.google.com/file/d/1srRuppK65BzD58EdKf6FkOpIGy5chI3-/view?usp=drive_link",
    "143": "https://drive.google.com/file/d/1XJANV2IJI0wwUjuhovFERPzA7yHLrwrM/view?usp=drive_link",
    "144": "https://drive.google.com/file/d/1cc2eFNjijEoyFlWqHMS4LiBLUH7IJGJb/view?usp=drive_link",
    "145": "https://drive.google.com/file/d/1GDk8EqBpHBIwKwnN3xa2YA7vZTq3uG36/view?usp=drive_link",
    "146": "https://drive.google.com/file/d/1njiMw45Soa1wjp4UAUFwCVD52BzOlFFJ/view?usp=drive_link",
    "147": "https://drive.google.com/file/d/1GnhixLIIMYvn95J6Fsy_zB9Buk-aE0Rv/view?usp=drive_link",
    "148": "https://drive.google.com/file/d/1ZWsDs_WZwgF5PbTwptBuuwfT3WU7UHBv/view?usp=drive_link",
    "149": "https://drive.google.com/file/d/1H-0ksuYEp-TAm5BT4rq1NM9MkJEuX5XT/view?usp=drive_link",
    "150": "https://drive.google.com/file/d/1tSRzIpUO_TubWq-nJY4RUCJyXkCISX5t/view?usp=drive_link",
    "151": "https://drive.google.com/file/d/1kiaEH-Swbw4MuDlktkldLLdfnmahRSJ0/view?usp=drive_link",
    "152": "https://drive.google.com/file/d/1kimS52E-O_Gk_731fPVw_yfUmtn8Xj-9/view?usp=drive_link",
    "153": "https://drive.google.com/file/d/1SPhMbGMKEfnZ6KKrl9fO0VhU6_NFJhCE/view?usp=drive_link",
    "154": "https://drive.google.com/file/d/1N6LwRhr8VvgnYcBrUwvn268j9AXdTb8C/view?usp=drive_link",
    "155": "https://drive.google.com/file/d/188z6LuAd4HHf_GpMNaFZ4oiIqoOcY7YT/view?usp=drive_link",
    "156": "https://drive.google.com/file/d/1yKNUYJ2vPcNzvrtaojtS09EdLTPbqbVZ/view?usp=drive_link",
    "157": "https://drive.google.com/file/d/1909JAYG5jfVWBblpbgBQGAZo2fDx2lHk/view?usp=drive_link",
    "158": "https://drive.google.com/file/d/1v1k6PmMEO91ULNBrB8yPvb5Hb1dLHTK1/view?usp=drive_link",
    "159": "https://drive.google.com/file/d/1C5yNzdhs2m8Itio_PQ5dl08yc5_zs3XP/view?usp=drive_link",
    "160": "https://drive.google.com/file/d/1m7WPGQOO6s_zpGb3Vctf_81q585hWjH7/view?usp=drive_link",
    "161": "https://drive.google.com/file/d/1IqHkYIB7gHKUqrlIk0-qO4Piovj5NKvd/view?usp=drive_link",
    "162": "https://drive.google.com/file/d/1v1qdENWzaZ9EoGvPCL7Nb9-8y5cq8gux/view?usp=drive_link",
    "163": "https://drive.google.com/file/d/1UyF9xTJpm2I32XKmwc9ZbSrC_wjWboWu/view?usp=drive_link",
    "164": "https://drive.google.com/file/d/1BXjBPO_v1hAiH5jtEq81clMaJYBdoliI/view?usp=drive_link",
    "165": "https://drive.google.com/file/d/1dhWxoxa_h3xsKql1gBvB7NqYZqx9e0ug/view?usp=drive_link",
    "166": "https://drive.google.com/file/d/1d-L-6U8kB05WyyufJlA_V85O3pp9u3Tg/view?usp=drive_link",
    "167": "https://drive.google.com/file/d/1gt2BmtzHu4oxcXPkxSOSp6KZI1BKTHww/view?usp=drive_link",
    "168": "https://drive.google.com/file/d/1YAJscHTRY0TbyIgCb2X-0Nv-VG5sdqmp/view?usp=drive_link",
    "169": "https://drive.google.com/file/d/14J8cS5OEn8n5Jkzb90_OaUk4ksUd8Ow_/view?usp=drive_link",
    "170": "https://drive.google.com/file/d/1-_c-R0Az3R18D1m6dNQCR16Ah7WxS2_U/view?usp=drive_link",
    "172": "https://drive.google.com/file/d/169on_foa52vkMiWuyDF6UMwGNCazh9Lv/view?usp=drive_link",
    "173": "https://drive.google.com/file/d/1OEOr7jf9wVv8GLmB9LT_LNYXdt37PF11/view?usp=drive_link",
    "174": "https://drive.google.com/file/d/1Tlqd454LBwGCx3WMIpLoq-R7u3hJEFlA/view?usp=drive_link",
    "175": "https://drive.google.com/file/d/1VzfNCRyPW5goj5fvfc88wJka4x_m95Di/view?usp=drive_link",
    "176": "https://drive.google.com/file/d/1IYJN4pqf3xt41m38NKhq1Qhw5JWeajdP/view?usp=drive_link",
    "177": "https://drive.google.com/file/d/1hsVu3X3jppdE0bZk4Wpm0JbXjA6dKCD-/view?usp=drive_link",
    "178": "https://drive.google.com/file/d/1XzABI1MrGPV3E7u6valp753vZaNYB_HD/view?usp=drive_link",
    "180": "https://drive.google.com/file/d/1dyJKxp6aRSa2gNZkSSwTv-xOpPYY2bgw/view?usp=drive_link",
    "181": "https://drive.google.com/file/d/1aAmxYQRE2UpXXwb5UpXTN8QVNX2CAHXF/view?usp=drive_link",
    "182": "https://drive.google.com/file/d/1XFARmOg6lYhOXNIlHu5WK3PvM6qTaAZE/view?usp=drive_link",
    "183": "https://drive.google.com/file/d/1M759qtK5PYxpCUDsh13rqO2LYROmOKsy/view?usp=drive_link",
    "184": "https://drive.google.com/file/d/1S5qI4kZOOJUkVzmH33DLpDryCLMhAGOy/view?usp=drive_link",
    "185": "https://drive.google.com/file/d/1oLKCM3O7zQTl1QI6Su-QfB9J3Qr7VybY/view?usp=drive_link",
    "186": "https://drive.google.com/file/d/1ha8H8AQpdrwwyvdEA-O2ZxBxb2TidJzV/view?usp=drive_link",
    "187": "https://drive.google.com/file/d/1wOzSfKqPy9uUta5LBZHWxV8DF2vLJTIL/view?usp=drive_link",
    "188": "https://drive.google.com/file/d/1rbQBQxZEO8U92q0Nqa6NYuKURzbBBy-A/view?usp=drive_link",
    "189": "https://drive.google.com/file/d/1E-GWe4OXMtNoWzLAiIOlGxIydzjAqKiW/view?usp=drive_link",
    "190": "https://drive.google.com/file/d/12ZyPsQKHjLvMXjHans177dh7080py3JO/view?usp=drive_link",
    "191": "https://drive.google.com/file/d/1y2vttRQRdE-G58NYZghcEoPVLykc911L/view?usp=drive_link",
    "193": "https://drive.google.com/file/d/19mOtsJtZiINcHFUd0l1KoG_IeEqc6HGr/view?usp=drive_link",
    "195": "https://drive.google.com/file/d/1Ekgt_Zt_kjxwnlWjgk6OxrJceVQRvOgy/view?usp=drive_link",
    "196": "https://drive.google.com/file/d/1niHDuD16ZkcMCVffM0VbNSBM2TZRgZSN/view?usp=drive_link",
    "197": "https://drive.google.com/file/d/1LqGc1nWa0Ae_h87iVnofhAHp8F5bc0WX/view?usp=drive_link",
    "198": "https://drive.google.com/file/d/1Zz0vpCmn2ljfh6CUm4Gz5x_n8HBh5TJV/view?usp=drive_link",
    "199": "https://drive.google.com/file/d/18beh9O3y_ca-JkUNkFj4XMtxHBh7GwC3/view?usp=drive_link",
    "200": "https://drive.google.com/file/d/13LJP1ehR2aoQaaznB6y3IfzzHV-zHF2-/view?usp=drive_link",
    "202": "https://drive.google.com/file/d/1NF-3sVL5bnii50tFIROjdh6z7Lo76USU/view?usp=drive_link",
    "203": "https://drive.google.com/file/d/10-M1J05tsUfExC9hMFza6kWMMXaw63_D/view?usp=drive_link",
    "204": "https://drive.google.com/file/d/15VPmvoQ4BrJUwOqdbLv7I68O1fmHHc2r/view?usp=drive_link",
    "205": "https://drive.google.com/file/d/17dOKUiQMoLFJb36rPn2BCLjp34EWc4Cq/view?usp=drive_link",
    "206": "https://drive.google.com/file/d/1UfOo3wKa-eAq6EezcZGv2-3nyzJKveMP/view?usp=drive_link",
    "207": "https://drive.google.com/file/d/1vRx1Db2MTOkQCs80p3nHx9sLfD1wJ-Hf/view?usp=drive_link",
    "208": "https://drive.google.com/file/d/1FxGlWcIA60FG2GbTn_DFDgqqjPJ58cT0/view?usp=drive_link",
    "209": "https://drive.google.com/file/d/1dooEGrb5gdrtRA8JOUHxUk-n5a7XsjYq/view?usp=drive_link",
    "210": "https://drive.google.com/file/d/1w5sXH7BHB01EEu8OM4gkg3oxeHG9p4Iq/view?usp=drive_link",
    "211": "https://drive.google.com/file/d/1hOn6etwkEHugHYuA_cfpI9xR_MNluKtE/view?usp=drive_link",
    "212": "https://drive.google.com/file/d/1afonMSD2-AHlB43EjE7YspCClWk-PsZL/view?usp=drive_link",
    "213": "https://drive.google.com/file/d/1gNjWgZ2WyUhd1cYZ2FdVrwRDnsjcb5O-/view?usp=drive_link",
    "214": "https://drive.google.com/file/d/130RyyGbVUBsBT4Y1hnpvF02VDvqASTOP/view?usp=drive_link",
    "215": "https://drive.google.com/file/d/19ish-hKHA1ZrDZlDrRtCJSIKtBG44916/view?usp=drive_link",
    "216": "https://drive.google.com/file/d/1rErP-N0QciDODFBzaKCM1q_6MnOFtCi_/view?usp=drive_link",
    "217": "https://drive.google.com/file/d/1p2aLSewySKI0oAogcc9hbChFlkwsZ1TT/view?usp=drive_link",
    "218": "https://drive.google.com/file/d/1Nt9LUBndNwl0HuaTRGyVh4Tt8L8myvwT/view?usp=drive_link",
    "219": "https://drive.google.com/file/d/14C45DhPjkk6gAceOOxw8S4MPwcUvyWkn/view?usp=drive_link",
    "220": "https://drive.google.com/file/d/1vIkxwr4heDZAh2wbTVVQN5khTb2obYPn/view?usp=drive_link",
    "221": "https://drive.google.com/file/d/1DFtE1UA2TVeKp2gSWDsd0Ay0NN_wrJoP/view?usp=drive_link",
    "222": "https://drive.google.com/file/d/1AbzoiQHVf__-WYR4oRflAAaomOK1Dk4h/view?usp=drive_link",
    "223": "https://drive.google.com/file/d/1SIAMlk2Ze9HiQN_QKNHdCBI8g5doXQUt/view?usp=drive_link",
    "224": "https://drive.google.com/file/d/17q8xJBxQnEkNn-72OhJzs2Rxy2Ko7g8K/view?usp=drive_link",
    "225": "https://drive.google.com/file/d/1nRMkSiyD5J4PbV15GPY7sII54jhTk63Z/view?usp=drive_link",
    "226": "https://drive.google.com/file/d/1PoGMGqXYd34ethz2Mcge6N1uvPuNnrNR/view?usp=drive_link",
    "227": "https://drive.google.com/file/d/12sqOw9tKNDG-ine6gFLgKBCdM9NBRqfX/view?usp=drive_link",
    "229": "https://drive.google.com/file/d/1SNYcukRAyyo3uxBIczaeR5LyubcvPbmx/view?usp=drive_link",
    "230": "https://drive.google.com/file/d/1KGj3XQ9ncBaAOnig8JmjLcoBqWlg_1p8/view?usp=drive_link",
    "231": "https://drive.google.com/file/d/1aVkfIFnygy_VyNJdOrlAo66XXMbNVfLW/view?usp=drive_link",
    "232": "https://drive.google.com/file/d/1CyR263F4Y8X3UPpUbRSii3p2uKlSfFNl/view?usp=drive_link",
    "233": "https://drive.google.com/file/d/1Cvnb0Akefmm71RBkCU7CMLNbbki-v1XS/view?usp=drive_link",
    "234": "https://drive.google.com/file/d/1u0b3pXupmXKLzwuI_Dw-03FRTMch6xL_/view?usp=drive_link",
    "235": "https://drive.google.com/file/d/1xfUY3zmH4EeNAy-yC20Cuo-aLA7LAsXe/view?usp=drive_link",
    "236": "https://drive.google.com/file/d/1yw0uX84UC1Wj9Ta_7S-7GYd6V2rVvLZj/view?usp=drive_link",
    "237": "https://drive.google.com/file/d/1M1SqBGn6HVo86S0dKiaW8nIaWZT2NqjF/view?usp=drive_link",
    "238": "https://drive.google.com/file/d/10Zf2YZpviBvlbWkI6vxpGSleVqNlOJRE/view?usp=drive_link",
    "239": "https://drive.google.com/file/d/1NcCA6jz9rrxiZK8EaGs7xRjCf8Ln-1F3/view?usp=drive_link",
    "240": "https://drive.google.com/file/d/1nPPYDrPF3_j_PUxfHtcRQxMPEUFHlGVG/view?usp=drive_link",
    "241": "https://drive.google.com/file/d/1THSwmVyFVultxCkh7nEtpEe9Wdf76Ohb/view?usp=drive_link",
    "242": "https://drive.google.com/file/d/1OMl-oZZrsRzj_rF78boiCGqf41GBXgqm/view?usp=drive_link",
    "243": "https://drive.google.com/file/d/1hDCkeH-BIw5IlwOKvz5nNjPYwue9TL_t/view?usp=drive_link",
    "244": "https://drive.google.com/file/d/1nkOIE4PpiXB_i2Az2pfqQt03GPULih2V/view?usp=drive_link",
    "245": "https://drive.google.com/file/d/1nWaEwrBsUcSx2adKcFZPxUp9-xxXzOQX/view?usp=drive_link",
    "246": "https://drive.google.com/file/d/1ihKf9xeS_6eG5iQn6o7CwNpyeCRJkoDN/view?usp=drive_link",
    "247": "https://drive.google.com/file/d/1HuBHTG0C1-BMGl6HkYE2V5Nru7h9h6em/view?usp=drive_link",
    "248": "https://drive.google.com/file/d/16cW9CMCowfE3u2EUQ1-64PgymnXSq1-R/view?usp=drive_link",
    "249": "https://drive.google.com/file/d/1LfXOYRNKMhD_TaTpqjKV_3pvNsS9vBB5/view?usp=drive_link",
    "250": "https://drive.google.com/file/d/1uFRX8SKw2906blw0j8zIzmHPGoYbChuK/view?usp=drive_link",
    "251": "https://drive.google.com/file/d/1b0vMPGAuWKUqCR6af6NSRNTDTPpGZNtb/view?usp=drive_link",
    "252": "https://drive.google.com/file/d/1xyRi1k3aWIDXsDCpYhF0XGqk6SN4FVTq/view?usp=drive_link",
    "253": "https://drive.google.com/file/d/11kXDOhVDJkQDZf63arOMZ-7b_ORUijRE/view?usp=drive_link",
    "254": "https://drive.google.com/file/d/1v_h5XrBYmvjPHA964WRckSUaZSLCI5nq/view?usp=drive_link",
    "255": "https://drive.google.com/file/d/1jf6mr14togYLBpQmdoffJTXVpOBxcTzn/view?usp=drive_link",
    "256": "https://drive.google.com/file/d/1HjzHcm-GNqHXLYVRILXW7DHv91BH6l3s/view?usp=drive_link",
    "257": "https://drive.google.com/file/d/1QnzI9wAVIFcy3TTVn7t10sAqEp9V-3wB/view?usp=drive_link",
    "258": "https://drive.google.com/file/d/1zQI6koGZEwrl_DDZMwI6TuHGC4BJfzUL/view?usp=drive_link",
    "259": "https://drive.google.com/file/d/1aIzcImxmBdYnzBCDG-0WSrv7b_bSW2hw/view?usp=drive_link",
    "260": "https://drive.google.com/file/d/1LXoz8BB8P4zfZXqkKQ8TluxK0X79pCqD/view?usp=drive_link",
    "262": "https://drive.google.com/file/d/1MYkJRqI-RpoX_i3wm9zGBjXLV2K4hdZg/view?usp=drive_link",
    "263": "https://drive.google.com/file/d/100M9lWxElniK4nbh2Y0ArMe4OeYXMUbn/view?usp=drive_link",
    "264": "https://drive.google.com/file/d/1iP4ah9D1snSjfPc3HjfpcUDYB8Ydig1v/view?usp=drive_link",
    "265": "https://drive.google.com/file/d/1uZHCGY1WM9RsygdbHt3_IbogqzXDbNe-/view?usp=drive_link",
    "266": "https://drive.google.com/file/d/1Df20bZF28ugcG0k6jIvaI8huScTYrtBt/view?usp=drive_link",
    "267": "https://drive.google.com/file/d/1xa1hBLI-ke4PEu2G52sdJdDd_SVUIFZV/view?usp=drive_link",
    "268": "https://drive.google.com/file/d/15MfLFIX1-4RghJYf0Xjj-nnsxImGddUj/view?usp=drive_link",
    "270": "https://drive.google.com/file/d/12iqjnLcXo6ExcgMLgZR9BvCLfNtV-Sia/view?usp=drive_link",
    "271": "https://drive.google.com/file/d/1j1qtOXb1FsLmGkhK75BTm-QVO0bQF3Jt/view?usp=drive_link",
    "272": "https://drive.google.com/file/d/1ENpXW1wYfxq7EyUElfUAvFRPryR08Dsd/view?usp=drive_link",
    "273": "https://drive.google.com/file/d/14O0ze7h0quiJQ-laSrT1-IEPy09PNNMC/view?usp=drive_link",
    "274": "https://drive.google.com/file/d/11LWRj9QfOrGptmvcOAsnp2Bw7Pbf01HD/view?usp=drive_link",
    "275": "https://drive.google.com/file/d/1Kku1rIDf1Jgr5PKNnnG4cyuZoQsX14vL/view?usp=drive_link",
    "276": "https://drive.google.com/file/d/1yBiv65Ljf5uQIk00TwfC-UHeCQ12SaB5/view?usp=drive_link",
    "277": "https://drive.google.com/file/d/1B0dJQPPyi9xTG4wKv1QXrCa5FAZ49Gl4/view?usp=drive_link",
    "278": "https://drive.google.com/file/d/1rZKsy3ConVEojkLXxqcLJeDiU7FV9FlT/view?usp=drive_link",
    "279": "https://drive.google.com/file/d/17z4PlSis3AlVylaCWiOTH6RtEDlRb5Gx/view?usp=drive_link",
    "280": "https://drive.google.com/file/d/1sbHdXeI_8pLXTOD4egcBYoZqF3Iz7hFL/view?usp=drive_link",
    "283": "https://drive.google.com/file/d/1waFOSCTvO8j98x1aWs7IKRKdYq3ClO86/view?usp=drive_link",
    "284": "https://drive.google.com/file/d/1ywAgCrSgiRVSoGfKOVEw93MSv0r2mApM/view?usp=drive_link",
    "285": "https://drive.google.com/file/d/1NKwLym_Zj3U7XzzSNnoAc0ZOV2ztGKCq/view?usp=drive_link",
    "286": "https://drive.google.com/file/d/1xsK_uDQMAbIgnxgt-9P27P7_-tnmEtPV/view?usp=drive_link",
    "287": "https://drive.google.com/file/d/1Ia2oP1yLIJosBVfiJx6bcAn2wKtBJOKC/view?usp=drive_link",
    "289": "https://drive.google.com/file/d/1Q9C1N1OQNvNkIXeL01iLMqds1fvnsGAM/view?usp=drive_link",
    "290": "https://drive.google.com/file/d/1vepMuXexDmtF0WQ4sKfBGEkkYQP0qkqA/view?usp=drive_link",
    "291": "https://drive.google.com/file/d/11z6LVkwfynm_ztcz2Hm-YRQ8eAovQy-E/view?usp=drive_link",
    "292": "https://drive.google.com/file/d/1M9I33Ma9x4-_ss2G0y6Z4yVnPZB-fK8v/view?usp=drive_link",
    "293": "https://drive.google.com/file/d/1z3fWTh6WH6EnJUYQKBZeNKMEHXsEMv12/view?usp=drive_link",
    "295": "https://drive.google.com/file/d/1YELTdjtzzrLHA0k3JZpJZjBtCBIm_n2y/view?usp=drive_link",
    "296": "https://drive.google.com/file/d/11lMLQp-10qxRtkAjDOFwud3WjSSc2JVm/view?usp=drive_link",
    "297": "https://drive.google.com/file/d/1gDw2CUNyChfn30MdhYbLjq2uQgurQPv5/view?usp=drive_link",
    "298": "https://drive.google.com/file/d/1XEH93Gw1kiPMCBXlEzEkhLx9jiQtLEcF/view?usp=drive_link",
    "300": "https://drive.google.com/file/d/1cfmituaGY8ZBgpsVzCAqNurZi3fPzGgS/view?usp=drive_link",
    "301": "https://drive.google.com/file/d/1rUa9VxtxAUEsGcL2ZxsyCTEqXJXFHxgM/view?usp=drive_link",
    "302": "https://drive.google.com/file/d/1UPArKkwQ5UtyES6iqGIWdwWC9T821VOC/view?usp=drive_link",
    "303": "https://drive.google.com/file/d/1NmCJ-37aDEJeS9lVTn-Pn4OnuU_lE9pm/view?usp=drive_link",
    "304": "https://drive.google.com/file/d/1TEpGak_Sb5lWKclM5oH_ErcyzcCNH87s/view?usp=drive_link",
    "305": "https://drive.google.com/file/d/1iVIMRhadrzxXL5gXZavVn0Psn6wCufqA/view?usp=drive_link",
    "306": "https://drive.google.com/file/d/18Hs3Q7BBAoyiSa6NLTP_HciAnmPwzlcj/view?usp=drive_link",
    "307": "https://drive.google.com/file/d/18TgtzKSpRw5IFNtSJB5HJOD53nmDS5qD/view?usp=drive_link",
    "308": "https://drive.google.com/file/d/14tAGiUVgMInt9Hlv3u1hWtMeDhzFy9Qk/view?usp=drive_link",
    "309": "https://drive.google.com/file/d/1m6ocexPF54YkNtKDaAsw6a6xRSf07XD_/view?usp=drive_link",
    "310": "https://drive.google.com/file/d/1M7WlCyiVqNf_FmmlpPzhvQIIUVFK3sKn/view?usp=drive_link",
    "312": "https://drive.google.com/file/d/1LhTanPVVzxp1e7d34YJ-wcqGCQM7l6XZ/view?usp=drive_link",
    "313": "https://drive.google.com/file/d/1t0XGXkYmY0kWh_wcBxWlfkLa0dY7bWAE/view?usp=drive_link",
    "314": "https://drive.google.com/file/d/1GgJ7s9sYF5uXQnn_2LWbHUN2CkcEWUid/view?usp=drive_link",
    "315": "https://drive.google.com/file/d/11lJdgsuHbNTQYTB-AxOkKyFi3X8lulGI/view?usp=drive_link",
    "316": "https://drive.google.com/file/d/1oBhUiUa01B_eUyI3p9czTqp56SB-jDoO/view?usp=drive_link",
    "317": "https://drive.google.com/file/d/1fOln-U6rWSDHfWkIBflKMhlwX6eM5YVI/view?usp=drive_link",
    "318": "https://drive.google.com/file/d/1s7-2urkRNNHD9HcbBq3qKS-7GpNmysEW/view?usp=drive_link",
    "319": "https://drive.google.com/file/d/1QBPPa3kw5tKbE8hts-yiaDAdWOnNKJ6H/view?usp=drive_link",
    "320": "https://drive.google.com/file/d/1nMEc0rIUCoDNNlqIE4ZYb8-z_vd9hdzA/view?usp=drive_link",
    "323": "https://drive.google.com/file/d/1F-7AalZXLJ6HJq8DwcybxJ4pu7wwq8uA/view?usp=drive_link",
    "324": "https://drive.google.com/file/d/1i757rYVXpMDZqSb-123MqlXMu9LuVWQv/view?usp=drive_link",
    "325": "https://drive.google.com/file/d/1WDz6Te0BnLMlHrhbLUIIBv3W1eNu71f4/view?usp=drive_link",
    "328": "https://drive.google.com/file/d/1VUCeB7HpMBNf0mtNExxiG9fnoH73AOUG/view?usp=drive_link",
    "329": "https://drive.google.com/file/d/1EK0zGQNvEzoQvOuqgjl8nwZvpWYGU9EM/view?usp=drive_link",
    "330": "https://drive.google.com/file/d/1Fwpog5RDBO_hLH7ByIBpHWKCIAAXi4XV/view?usp=drive_link",
    "331": "https://drive.google.com/file/d/1GCNjB3u8PIHPpj9c4CznP6UoiSbh9glA/view?usp=drive_link",
    "332": "https://drive.google.com/file/d/1VETPyoeAOQd7_C8CS5QwgxH8gdve53Ks/view?usp=drive_link",
    "333": "https://drive.google.com/file/d/1Rrwjj8GOW8bFVa6-SLkHmt3WQjMXMIcg/view?usp=drive_link",
    "334": "https://drive.google.com/file/d/1ZaRbSW_nOqrG04GrwKJ3CT0c1AoSCrD0/view?usp=drive_link",
    "335": "https://drive.google.com/file/d/1Owwno8BBv4yqTej9HIteeNdnTYBtdVAj/view?usp=drive_link",
    "336": "https://drive.google.com/file/d/1gMjgfqTC1upFDgYostuYzY8VZ_g9J9aQ/view?usp=drive_link",
    "337": "https://drive.google.com/file/d/1kLJ2w6fXtmJ20a7vnDRvfdzyg72G3npN/view?usp=drive_link",
    "338": "https://drive.google.com/file/d/1_wh0YZ2v6Y5R78aSHPxXl8nOfCwYIYDB/view?usp=drive_link",
    "339": "https://drive.google.com/file/d/1VO_isuLAaRLiGZSRs8xN1EmU3rSmwyS-/view?usp=drive_link",
    "340": "https://drive.google.com/file/d/17HFmPDaiYhd08PSzYHr9-qz0jfAMVzEQ/view?usp=drive_link",
    "341": "https://drive.google.com/file/d/1vwijtKtdilnsG5-g6yIBcH_q6dhZR5hl/view?usp=drive_link",
    "343": "https://drive.google.com/file/d/148kdxZj2a_QA-PGfIco7-w_4MYTfa09R/view?usp=drive_link",
    "345": "https://drive.google.com/file/d/1oA7XXldrjqcnHm7ZQVgsxL-SjU75n4XW/view?usp=drive_link",
    "347": "https://drive.google.com/file/d/14i9V0UoU7KRJUqcFJ-77J3fkE3zMFvwF/view?usp=drive_link",
    "348": "https://drive.google.com/file/d/1i16qVyfZ5zwBb6GX9OQhHsXuBMcQ0cQH/view?usp=drive_link",
    "349": "https://drive.google.com/file/d/104oP-JzLoOgD0PDuSe80Mc4rWymdTqxt/view?usp=drive_link",
    "350": "https://drive.google.com/file/d/1UahGkCgJgrMHiGDhGd_bpC2rYTAOS60V/view?usp=drive_link",
    "353": "https://drive.google.com/file/d/1QJEmq3U46aaSNuiL-1kSRzP6L268MNhJ/view?usp=drive_link",
    "355": "https://drive.google.com/file/d/1f2PlpzqWAn9NYK7qR5dNfy97DqocKkhi/view?usp=drive_link",
    "356": "https://drive.google.com/file/d/1CKT8yZnL6tZQpAeazkoeJdLluy04Bf1G/view?usp=drive_link",
    "357": "https://drive.google.com/file/d/1SZHl4nX38ZPSrq8jYJJVblDWS27JPQBQ/view?usp=drive_link",
    "358": "https://drive.google.com/file/d/1SO06BabGKqy7mNL7j82gHtBxHKev5Vop/view?usp=drive_link",
    "359": "https://drive.google.com/file/d/1ABLybHrZFughODVhw5_iDYcMzmosPaAQ/view?usp=drive_link",
    "360": "https://drive.google.com/file/d/12kwQUqc2fNS3QwaPbCXnrPq4OcFgRrnV/view?usp=drive_link",
    "361": "https://drive.google.com/file/d/1YIbEvsdPNtn_LswXKvKxiV17oBiV_tWr/view?usp=drive_link",
    "362": "https://drive.google.com/file/d/1qPu5bhCHB1VyGySTilVJBnyOoMm-Uvo2/view?usp=drive_link",
    "363": "https://drive.google.com/file/d/1tqBnHutbRE_X8R8FKv-yRvLKCZJKs3nO/view?usp=drive_link",
    "364": "https://drive.google.com/file/d/152j3jWw__lnLuXGCm7nYMvjLmXLz0-ci/view?usp=drive_link",
    "365": "https://drive.google.com/file/d/11LVzOWZL3IxRQsAOWiZNMGSr1BTYobB9/view?usp=drive_link",
    "366": "https://drive.google.com/file/d/1_Z5kyTSF9t23oGbeBB74wKGxvHrcYSVb/view?usp=drive_link",
    "367": "https://drive.google.com/file/d/1u8kE_EOpt3RcUoAAH_8YJnLlfUcli-l-/view?usp=drive_link",
    "368": "https://drive.google.com/file/d/1uW0xcwURbh5Rop4WWZPi_9o_cxWA8eMI/view?usp=drive_link",
    "369": "https://drive.google.com/file/d/1-dQWZjh8NLT_cCoZJJ3VuUogN9ifDpA-/view?usp=drive_link",
    "370": "https://drive.google.com/file/d/10TSR4aY1GWlPsWiOZc253TKlsh9rHR4I/view?usp=drive_link",
    "371": "https://drive.google.com/file/d/1uJ4eIL6M_MUGpQ2_UUorL0zkpaN5m454/view?usp=drive_link",
    "372": "https://drive.google.com/file/d/1mgOuBNJyIkHduujAhQFRNK4iBQJnUTTH/view?usp=drive_link",
    "373": "https://drive.google.com/file/d/1hFYJb1EmafnPUI5AL6Up8as0fcL3bSUA/view?usp=drive_link",
    "374": "https://drive.google.com/file/d/1BU36cqLqOMy3_jqkxGmqFa2UV8P68hUv/view?usp=drive_link",
    "375": "https://drive.google.com/file/d/1Dr7C2OIzkkMJnv8wLY6LraQWYD1GGzIg/view?usp=drive_link",
    "376": "https://drive.google.com/file/d/1KH78dqjH6ak_FpMkY7Uzqqf2GSagiLdK/view?usp=drive_link",
    "377": "https://drive.google.com/file/d/1qIL80JGramcG02WzAuGkJayJrLiDCVIL/view?usp=drive_link",
    "379": "https://drive.google.com/file/d/13YrsLGoi_dSUPARNeG74F2A7mKmaha64/view?usp=drive_link",
    "380": "https://drive.google.com/file/d/1R7u66kQSdEK1dXxtP3GYhD_iziorGrYX/view?usp=drive_link",
    "381": "https://drive.google.com/file/d/1XSgdPD_-GRQDuAmsT_zP295RpkyInPLq/view?usp=drive_link",
    "383": "https://drive.google.com/file/d/1LMkFgkCyuL7wG1P266d98DIv5fiIplZ6/view?usp=drive_link",
    "384": "https://drive.google.com/file/d/1jM5B1ikEy1HdMTtie0tsBKQ52c6eWQ0f/view?usp=drive_link",
    "385": "https://drive.google.com/file/d/1TD7bPlf1JbqTjMPHvbOsQyxMUla4VYyi/view?usp=drive_link",
    "386": "https://drive.google.com/file/d/1lFTsRUFW4RR9j0bB7d1B6lEzvUzto83j/view?usp=drive_link",
    "387": "https://drive.google.com/file/d/18RSKhLl5CezE4W97akPXZ7nFjfj8sDVy/view?usp=drive_link",
    "389": "https://drive.google.com/file/d/1HwOL9IOSJGfAa5tAy1rbP4s2zwRJz-qG/view?usp=drive_link",
    "390": "https://drive.google.com/file/d/1howh2nTbIG2TqHYbbsidnBgZK3axJ6QO/view?usp=drive_link",
    "391": "https://drive.google.com/file/d/1aR4tHakkdCbzaANhd8D-aLd7MNrJhY3h/view?usp=drive_link",
    "392": "https://drive.google.com/file/d/1rH5cUEEWe1ibNeYM-tBBvIEiD-1euZWr/view?usp=drive_link",
    "393": "https://drive.google.com/file/d/11zs6b76wiRbuRRohQxaNs1C-qNEgF_Nw/view?usp=drive_link",
    "394": "https://drive.google.com/file/d/1WZN9SGHol_ELsQ6nzlLjvADVX3oI4vHL/view?usp=drive_link",
    "395": "https://drive.google.com/file/d/1iLIuxC8SN2jm6O_r-pByBAP5Z-8rbwJL/view?usp=drive_link",
    "396": "https://drive.google.com/file/d/1NL-hW3DufaS3404hKxY2lQJy2Macodvb/view?usp=drive_link",
    "397": "https://drive.google.com/file/d/19_01G02sdzMBzj53Iw-1l7A-s5mwuYEe/view?usp=drive_link",
    "398": "https://drive.google.com/file/d/13XCu9G5h_W_-gcfnk-6ChPTwWeAw4w0S/view?usp=drive_link",
    "399": "https://drive.google.com/file/d/1L3HG2hXgNfVNhPGTTGBpekpPNXWOQxAK/view?usp=drive_link",
    "400": "https://drive.google.com/file/d/1Q6zQqiXtEDmId-jqM1SgQvZiqKSpDQN_/view?usp=drive_link",
    "401": "https://drive.google.com/file/d/1sri2PN3-QSCxFzfJEhMqfp3UvdhqQfz2/view?usp=drive_link",
    "402": "https://drive.google.com/file/d/1S-P8N5cmpO9OsPvFqKotTzs8c5P1K4Kp/view?usp=drive_link",
    "403": "https://drive.google.com/file/d/1-tkxZzsFoZZq8gwWMKFwQhWhIhK2BYgR/view?usp=drive_link",
    "404": "https://drive.google.com/file/d/1zwKiVzWtMaFtI5ORbX1d6ulUnqntNNuz/view?usp=drive_link",
    "406": "https://drive.google.com/file/d/1YcSoTHNR3IiKznL1VqYJVFsyzvLVW-jk/view?usp=drive_link",
    "407": "https://drive.google.com/file/d/18JidCUqBiGvgt-p3T4GPvkK1zqcroKcu/view?usp=drive_link",
    "408": "https://drive.google.com/file/d/1NJofzRq0nmM8Kuk0EADFWWOp7syX66-w/view?usp=drive_link",
    "409": "https://drive.google.com/file/d/1780gwD5WvoVA6Rpbu3ucpLfGXDm-iQmF/view?usp=drive_link",
    "410": "https://drive.google.com/file/d/1Vkl5L-K2QSottk4H604UL3CC_69Qr5Sw/view?usp=drive_link",
    "411": "https://drive.google.com/file/d/1de-7zx6GYmvnj-ofxCv5tyUWZu9N84Ij/view?usp=drive_link",
    "412": "https://drive.google.com/file/d/1ZoNLREi6Uxxu3jGe9qSZER37Sunv3HtH/view?usp=drive_link",
    "413": "https://drive.google.com/file/d/1I4oHAUIl1Jq40ngazJ9nEbjgchuv5ISD/view?usp=drive_link",
    "414": "https://drive.google.com/file/d/1WDpfCua-4bnAoXg6i8vUjKxliN2sQG6h/view?usp=drive_link",
    "415": "https://drive.google.com/file/d/1UL2IZpxeaYJ7-M0OrhCzLZzBY2IeNXLp/view?usp=drive_link",
    "416": "https://drive.google.com/file/d/1IpEFa93U2JnNzAj9pIAVqiUBg2h1a2XA/view?usp=drive_link",
    "418": "https://drive.google.com/file/d/1vIO-TXcoipazWnQGJTmma3NNzNfkGcmo/view?usp=drive_link",
    "419": "https://drive.google.com/file/d/12L6mZWMYogeunnwO7_yr2t7Y3_MOtWXD/view?usp=drive_link",
    "420": "https://drive.google.com/file/d/1DWMlLEb6oSnjDVFQybOIRKxdclZMTzF2/view?usp=drive_link",
    "423": "https://drive.google.com/file/d/1nPOq71kzbT4qrymn_gpPjFlOhnevDHxD/view?usp=drive_link",
    "424": "https://drive.google.com/file/d/1Vxf7Btc1TLB7BycTDZk9CTT6yCeEpUXa/view?usp=drive_link",
    "425": "https://drive.google.com/file/d/1Ttwr8FsshCfxJC9HzJKsgyIZ4jnH_y0_/view?usp=drive_link",
    "426": "https://drive.google.com/file/d/1hQNCeE2WxmFrLkjm76_DgNaz3Wzzrwuz/view?usp=drive_link",
    "428": "https://drive.google.com/file/d/1KkAAVBClsJQPzFGd0bfBk329odNoU4Yo/view?usp=drive_link",
    "430": "https://drive.google.com/file/d/1XzIx3EGxb4Lf0VQw8kp0wL0uwBTmMLXE/view?usp=drive_link",
    "432": "https://drive.google.com/file/d/1w5BTsoB3jWHPYCKmZqKkjXj19EmRVmeh/view?usp=drive_link",
    "433": "https://drive.google.com/file/d/15QZDiJMWuIPktPcEaaYj4x4ioUUIEVZn/view?usp=drive_link",
    "434": "https://drive.google.com/file/d/1aK4zXUhIs7mtlJ1IYWBrlJCK2vbKZxGE/view?usp=drive_link",
    "437": "https://drive.google.com/file/d/11T_Y-_TnpOm_z0PpM2rFcGmDiSvy7vx3/view?usp=drive_link",
    "438": "https://drive.google.com/file/d/1WyPSRbb7_14pQ9walfWcCLt8i9B-brwE/view?usp=drive_link",
    "439": "https://drive.google.com/file/d/1fMyuEJU1ScwhY9HVyv1FCGQVIj5UKN47/view?usp=drive_link",
    "440": "https://drive.google.com/file/d/1T-2PFm3U3CBDH-JhFKiFnYG40vfrfkY4/view?usp=drive_link",
    "445": "https://drive.google.com/file/d/1B9zwad50oytAK5ZzUbJBxn4-MeNMWOq1/view?usp=drive_link",
    "446": "https://drive.google.com/file/d/1mlQf638XyTiu7-OqrQ7JHusd6cD2kUKM/view?usp=drive_link",
    "448": "https://drive.google.com/file/d/1y2rpHLCfk4lE8C6w0Sm7ubcW-NGXh906/view?usp=drive_link",
    "449": "https://drive.google.com/file/d/1YdnJ37YyHH-tSAq3DyelP8wLKNOGdJ0n/view?usp=drive_link",
    "450": "https://drive.google.com/file/d/1fTrf1sabnX81T4Dd-yhHSByemVuzbvnh/view?usp=drive_link",
    "451": "https://drive.google.com/file/d/1bzMtvaAOSYGBAmXcGWOk5j8plwweTEOO/view?usp=drive_link",
    "454": "https://drive.google.com/file/d/1bwvDWsmXlFPE73vE-KrtnLKL3elx45Uo/view?usp=drive_link"  # Aquí agregarías tus enlaces
    # ... más enlaces ...
}

enlaces_2025 = {
    "001": "https://drive.google.com/file/d/19tbduny-xI7gObZOs1Hx2EVsHeN-gh1i/view?usp=drive_link",
    "002": "https://drive.google.com/file/d/1nKty9CylYke621o49ssy35GRLf2-AyrX/view?usp=drive_link",
    "003": "https://drive.google.com/file/d/1lMdCAZ8XZe0X91gSt5QwMgz4P6L_-Jkx/view?usp=drive_link",
    "004": "https://drive.google.com/file/d/1OFsAtLd1g2UgoUMsnLEayhxRR59VWJd3/view?usp=drive_link",
    "005": "https://drive.google.com/file/d/1dbvniNT3HYzakwAWCeciqO5cZxzkUjg0/view?usp=drive_link",
    "006": "https://drive.google.com/file/d/1qkBwaAkhAKRLQojkPHELMBVxqQeWoGjU/view?usp=drive_link",
    "007": "https://drive.google.com/file/d/1xHlOttzcK2xJNSFgLzDnJkTrM0Zdl6lJ/view?usp=drive_link",
    "008": "https://drive.google.com/file/d/1sbf3ZGub0EnY9hY-zhVSz7NVkialOKfL/view?usp=drive_link",
    "009": "https://drive.google.com/file/d/1DvPtLbBKQL4UgchMKlLonH0Lq3JqaTSm/view?usp=drive_link",
    "010": "https://drive.google.com/file/d/1ZnBrlMaSrvLkwMJIZR6hrU1CDQfvNycV/view?usp=drive_link",
    "011": "https://drive.google.com/file/d/1_wrTwNe_Kj50bL8OJwAj1gCMp2xBsw1f/view?usp=drive_link",
    "012": "https://drive.google.com/file/d/13U7hLyT7bhBw9jaT6h0IdPQ-jJcl1rbh/view?usp=drive_link",
    "014": "https://drive.google.com/file/d/1ruJ-27vk9s-m45TwMeJv5NhaXo7hRtt5/view?usp=drive_link",
    "015": "https://drive.google.com/file/d/1fOlYosogyFRZ8V1jf33Fr8FwaYeucpmb/view?usp=drive_link",
    "016": "https://drive.google.com/file/d/1BdJcHqHRbWtePshg-WBZlGNtu3pC-Agm/view?usp=drive_link",
    "018": "https://drive.google.com/file/d/1yUdyzoIJ1PZ9j6x6F5G6xKyNnrvrYFlf/view?usp=drive_link",
    "020": "https://drive.google.com/file/d/1Oci6uLlU47c1nCwBtzZJfmq09zsbMPCh/view?usp=drive_link",
    "021": "https://drive.google.com/file/d/1VHkddCORq7SqFidBGFDHxTtNX6dzd_wq/view?usp=drive_link",
    "022": "https://drive.google.com/file/d/1coxzQyiqKpJlW8m_4HS0uZQtdUYvPW9d/view?usp=drive_link",
    "023": "https://drive.google.com/file/d/1lHJEBewYAH9qWpx1XIXRnIzXjl26L9xf/view?usp=drive_link",
    "024": "https://drive.google.com/file/d/1DPs4a1KdiMqbkLDaRHkrPfvD_RbUMrJm/view?usp=drive_link",
    "025": "https://drive.google.com/file/d/1sQX0qud9z901PTNEvEICEXYwJaSFItln/view?usp=drive_link",
    "028": "https://drive.google.com/file/d/1vpo-Xn7qOK_VJxQqYnPyRX2vFOIx3AOQ/view?usp=drive_link",
    "029": "https://drive.google.com/file/d/1CSnFDZUOsWULPHKdaY5e2zy-f3MpUFjk/view?usp=drive_link",
    "031": "https://drive.google.com/file/d/1DjSMwkPkAANTKt-jsWUnaJy5nYFLvili/view?usp=drive_link",
    "034": "https://drive.google.com/file/d/1k9ANxmGEYFG_IBbcXea4yeSEd3XSo0Mh/view?usp=drive_link",
    "038": "https://drive.google.com/file/d/12ZEVsrOOrRJt84nz41F2dQ2Vhm3ElYOX/view?usp=drive_link" # Aquí agregarías tus enlaces
    # ... más enlaces ...
}

# Función para crear enlace si existe el número de oficio
def create_link(row):
    try:
        num_oficio = row['NÚMERO OFICIO']  # Ya está en formato '001'
        if pd.isna(num_oficio):  # Si la celda está vacía
            return ''
        
        enlaces = enlaces_2024 if sheet_name == "Oficios 2024" else enlaces_2025
        
        if num_oficio in enlaces:
            return f'<a href="{enlaces[num_oficio]}" target="_blank" style="color: #1E88E5; text-decoration: none;">Ver Documento</a>'
        return ''
    except Exception as e:
        st.error(f"Error procesando fila: {e}")
        return ''

# Agregar columna de enlaces
df['ENLACE'] = df.apply(create_link, axis=1)

# Interfaz con Streamlit
st.markdown(f"<h1 class='stTitle'>Registro de {sheet_name}</h1>", unsafe_allow_html=True)

# Contenedor de filtros con diseño
with st.sidebar:
    st.header("Filtros 🔍")
    
    # Obtener valores únicos no nulos para los filtros
    asesor_valores = ["Todos"] + list(df["ASESOR"].dropna().unique())
    asunto_valores = ["Todos"] + list(df["ASUNTO"].dropna().unique())
    
    asesor = st.selectbox("Filtrar por Asesor", asesor_valores)
    asunto = st.selectbox("Filtrar por Asunto", asunto_valores)

# Aplicar filtros
df_filtered = df.copy()
if asesor != "Todos":
    df_filtered = df_filtered[df_filtered["ASESOR"] == asesor]
if asunto != "Todos":
    df_filtered = df_filtered[df_filtered["ASUNTO"] == asunto]

# Mostrar tabla con enlaces clickeables
st.write("""<div style='border-radius: 10px; background-color: white; padding: 20px;'>""", unsafe_allow_html=True)
st.write(df_filtered.to_html(escape=False, index=False), unsafe_allow_html=True)
st.write("""</div>""", unsafe_allow_html=True)