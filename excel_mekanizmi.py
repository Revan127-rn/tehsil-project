import pandas as pd

# Azərbaycan əlifbası ardıcıllığı
az_alphabet = {
    'a':'01', 'b':'02', 'c':'03', 'ç':'04', 'd':'05', 'e':'06', 'ə':'07', 
    'f':'08', 'g':'09', 'ğ':'10', 'h':'11', 'x':'12', 'ı':'13', 'i':'14', 
    'j':'15', 'k':'16', 'q':'17', 'l':'18', 'm':'19', 'n':'20', 'o':'21', 
    'ö':'22', 'p':'23', 'r':'24', 's':'25', 'ş':'26', 't':'27', 'u':'28', 
    'ü':'29', 'v':'30', 'y':'31', 'z':'32'
}

def az_sort_key(series):
    def map_string(s):
        return "".join([az_alphabet.get(char.lower(), char) for char in str(s)])
    return series.map(map_string)

def ballari_cixar(excel_fayli, sutun_adi):
    df = pd.read_excel(excel_fayli)
    ad_sutunu = df.columns[0]
    
    if sutun_adi not in df.columns:
        raise ValueError(f"Excel faylında '{sutun_adi}' başlığı tapılmadı.")
        
    df = df.dropna(subset=[ad_sutunu])
    
    def reqemi_temizle(val):
        val_str = str(val)
        if '/' in val_str:
            return val_str.split('/')[0].strip() # /3 hissesini atir, 45.5 qalir
        return val_str
        
    df[sutun_adi] = df[sutun_adi].apply(reqemi_temizle)
    
    # Əlifba ilə sıralayırıq
    df_sorted = df.sort_values(by=ad_sutunu, key=az_sort_key)
    
    # Sıralanmış yalnız balları siyahı kimi qaytarırıq
    return df_sorted[sutun_adi].tolist()
