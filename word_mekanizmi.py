import docx

def word_faylini_doldur(word_fayli, axtarilan_fenn, ballar_siyahisi):
    doc = docx.Document(word_fayli)
    # Şəkildəki əsas hesabat cədvəlini götürürük (sənəddəki ilk cədvəl)
    table = doc.tables[0]
    
    hedef_sutun_indeksi = -1
    baslangic_setir_indeksi = -1
    
    # Cədvəlin yuxarı hissəsindəki başlıqları axtarırıq
    for setir_idx in range(min(4, len(table.rows))):
        for sutun_idx, xana in enumerate(table.rows[setir_idx].cells):
            # Qısa ad yazılsa belə (məs: Rəqəmsal), tam adı tapacaq
            if axtarilan_fenn.lower() in xana.text.lower():
                hedef_sutun_indeksi = sutun_idx
                # Fənnin adından 2 sətir aşağıda tələbə məlumatları (B sütunu) başlayır
                baslangic_setir_indeksi = setir_idx + 2 
                break
        if hedef_sutun_indeksi != -1:
            break
            
    if hedef_sutun_indeksi == -1:
        raise ValueError(f"Word cədvəlində '{axtarilan_fenn}' başlığı tapılmadı.")
        
    # Balları ardıcıllıqla B sütununa yazırıq
    bal_idx = 0
    for setir_idx in range(baslangic_setir_indeksi, len(table.rows)):
        if bal_idx < len(ballar_siyahisi):
            # Hədəf sütun B xanasının dəqiq yerləşdiyi sütundur
            table.cell(setir_idx, hedef_sutun_indeksi).text = str(ballar_siyahisi[bal_idx])
            bal_idx += 1
            
    return doc
