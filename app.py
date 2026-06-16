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

            if not excel_fayli or not word_fayli:
                return "Hər iki fayl yüklənməlidir!", 400

            # 1. Excel-dən balları çıxarırıq (əlifba ilə düzülmüş və temizlenmiş)
            hazir_ballar = ballari_cixar(excel_fayli, excel_basliq)

            # 2. Word faylını yeniləyirik
            yenilenmis_word = word_faylini_doldur(word_fayli, word_fenn, hazir_ballar)

            # 3. Yalnız yeni Word faylını kompüterə yükləmək üçün hazırlayırıq
            output = io.BytesIO()
            yenilenmis_word.save(output)
            output.seek(0)

            return send_file(output, download_name="Yeni_Hesabat.docx", as_attachment=True)

        except Exception as e:
            return f"Sistem xətası baş verdi: {str(e)}", 500

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
