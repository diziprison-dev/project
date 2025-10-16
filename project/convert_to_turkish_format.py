#!/usr/bin/env python3
"""
Excel dosyasını Türkçe virgüllü formata çeviren script
"""
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from openpyxl.utils.dataframe import dataframe_to_rows

def convert_excel_to_turkish_format(input_file, output_file):
    """
    Excel dosyasını Türkçe virgüllü formata çevirir
    """
    try:
        # Excel dosyasını oku
        print(f"Excel dosyası okunuyor: {input_file}")
        df = pd.read_excel(input_file)
        
        print(f"Toplam {len(df)} satır yüklendi")
        print(f"Sütunlar: {list(df.columns)}")
        
        # Sayısal sütunları belirle
        numeric_columns = ['price1', 'price2', 'price3', 'price4', 'price5', 'buyingPrice', 'tax']
        
        # Sadece mevcut sütunları işle
        existing_numeric_columns = [col for col in numeric_columns if col in df.columns]
        print(f"İşlenecek sayısal sütunlar: {existing_numeric_columns}")
        
        # Yeni workbook oluştur
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Urun_Fiyatlari"
        
        # Türkçe sayı formatı oluştur
        turkish_number_style = NamedStyle(name="turkish_number")
        turkish_number_style.number_format = '#.##0,00'  # Türkçe format: 1.234,56
        
        # DataFrame'i worksheet'e ekle
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Başlık satırını kalın yap
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
        
        # Sayısal sütunları Türkçe formata çevir
        for col_name in existing_numeric_columns:
            col_idx = df.columns.get_loc(col_name) + 1  # +1 çünkü Excel 1'den başlar
            
            for row in range(2, len(df) + 2):  # 2'den başla çünkü 1. satır başlık
                cell = ws.cell(row=row, column=col_idx)
                
                # Sadece sayısal değerleri işle
                if pd.notna(df.iloc[row-2][col_name]):
                    try:
                        # Değeri float'a çevir
                        value = float(df.iloc[row-2][col_name])
                        cell.value = value
                        cell.number_format = '#.##0,00'  # Türkçe format
                    except (ValueError, TypeError):
                        # Sayısal olmayan değerleri olduğu gibi bırak
                        cell.value = df.iloc[row-2][col_name]
        
        # Dosyayı kaydet
        wb.save(output_file)
        print(f"Türkçe format dosyası kaydedildi: {output_file}")
        
        # Örnek değerleri göster
        print("\nÖrnek dönüştürülmüş değerler:")
        for col in existing_numeric_columns[:3]:  # İlk 3 sütunu göster
            print(f"\n{col} sütunu (ilk 5 değer):")
            for i in range(min(5, len(df))):
                value = df.iloc[i][col]
                if pd.notna(value):
                    print(f"  Satır {i+1}: {value} -> {value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        
        return True
        
    except Exception as e:
        print(f"Hata oluştu: {e}")
        return False

def main():
    input_file = "sturmmm.xlsx"
    output_file = "sturmmm_turkish_format.xlsx"
    
    print("Excel dosyasını Türkçe virgüllü formata çeviriyor...")
    print("=" * 50)
    
    success = convert_excel_to_turkish_format(input_file, output_file)
    
    if success:
        print("\n✅ Dönüştürme başarılı!")
        print(f"📁 Yeni dosya: {output_file}")
        print("\nBu dosyayı Excel'de açtığınızda sayılar Türkçe formatta görünecek:")
        print("Örnek: 1234.56 -> 1.234,56")
    else:
        print("\n❌ Dönüştürme başarısız!")

if __name__ == "__main__":
    main()