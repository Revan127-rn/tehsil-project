from flask import Flask, render_template, request, send_file
import io
from excel_mekanizmi import ballari_cixar
from word_mekanizmi import word_faylini_doldur

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            excel_fayli = request.files['excel_file']
            word_fayli = request.files['word_file']
            excel_basliq = request.form['excel_column']
            word_fenn = request.form['word_subject']
            
            # Formdan yeni faylın adını alırıq
            yeni_fayl_adi = request.form.get('output_filename', '').strip()
            
            # Əgər istifadəçi adı boş qoyarsa, standart ad veririk
            if not yeni_fayl_adi:
                yeni_fayl_adi = "Yeni_Hesabat"
                
            # Fayl adının sonuna .docx artırıldığından əmin oluruq
            if not yeni_fayl_adi.endswith('.docx'):
                yeni_fayl_adi += ".docx"

            if not excel_fayli or not word_fayli:
                return "Hər iki fayl yüklənməlidir!", 400

            # Məlumatları çıxarır və doldururuq
            hazir_ballar = ballari_cixar(excel_fayli, excel_basliq)
            yenilenmis_word = word_faylini_doldur(word_fayli, word_fenn, hazir_ballar)

            output = io.BytesIO()
            yenilenmis_word.save(output)
            output.seek(0)

            # download_name hissəsini istifadəçinin yazdığı ad ilə dəyişdik
            return send_file(output, download_name=yeni_fayl_adi, as_attachment=True)

        except Exception as e:
            return f"Sistem xətası baş verdi: {str(e)}", 500

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
