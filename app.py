from flask import Flask, render_template, request, send_file
import pandas as pd
import io

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 1. Veb saytdan gələn faylı və sütun adını götürürük
        file = request.files['excel_file']
        score_column = request.form['score_column']

        if not file or not score_column:
            return "Fayl və ya sütun adı boş ola bilməz!", 400

        try:
            # 2. Excel-i oxuyuruq
            df = pd.read_excel(file)
            
            # Adların olduğu birinci sütunu (A sütunu) müəyyən edirik
            name_col = df.columns[0] 

            if score_column not in df.columns:
                return f"Xəta: Sizin yüklədiyiniz faylda '{score_column}' adlı başlıq yoxdur. Zəhmət olmasa düzgün yazın.", 400

            # 3. Adı boş olan sətirləri (A2-dən başlayaraq) silirik
            df = df.dropna(subset=[name_col])

            # 4. Qiymət olan sütundakı məlumatı təmizləyirik (Məsələn: 45.5/3 -> 45.5)
            def clean_score(val):
                val_str = str(val)
                if '/' in val_str:
                    return val_str.split('/')[0].strip() # / isaresinden evvelki hisseni goturur
                return val_str

            df[score_column] = df[score_column].apply(clean_score)

            # 5. Adlara görə əlifba sırası ilə düzürük
            df_sorted = df.sort_values(by=name_col)

            # 6. Yalnız lazım olan 2 sütunu (Adlar və Təmizlənmiş Rəqəmlər) saxlayırıq
            df_final = df_sorted[[name_col, score_column]]

            # 7. Yeni faylı kompüterə yükləmək üçün hazırlayırıq
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Netice')
            output.seek(0)

            # Hazır faylı istifadəçiyə yüklədirik
            return send_file(output, download_name="Hazir_Siyahi.xlsx", as_attachment=True)

        except Exception as e:
            return f"Sistem xətası baş verdi: {str(e)}", 500

    # Əgər səhifəyə normal daxil olubsa, HTML dizaynı göstər
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)