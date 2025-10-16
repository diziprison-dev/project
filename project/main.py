import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import os

class ProductPriceCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Ürün Fiyat Hesaplayıcı - Product Price Calculator")
        self.root.geometry("1200x700")
        self.root.configure(bg='#f0f0f0')
        
        # Data storage
        self.df = None
        self.filtered_df = None
        self.selected_items = set()
        
        # Variables
        self.exchange_rate = tk.DoubleVar(value=1.0)
        self.coefficient = tk.DoubleVar(value=1.0)
        
        self.setup_ui()
        self.load_data()
        
        # Force focus on first entry after UI is ready
        self.root.after(100, self.force_initial_focus)
    
    def safe_float(self, value):
        """Convert Turkish number format to float safely"""
        if pd.isna(value) or value == "" or value is None:
            return 0.0
        try:
            # Convert to string first
            str_value = str(value).strip()
            # Replace Turkish decimal separator (,) with English (.)
            str_value = str_value.replace(',', '.')
            # Remove any non-numeric characters except decimal point and minus
            import re
            str_value = re.sub(r'[^\d.-]', '', str_value)
            # Handle empty string after cleaning
            if str_value == "" or str_value == "-":
                return 0.0
            return float(str_value)
        except (ValueError, TypeError):
            return 0.0
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Ürün Fiyat Hesaplayıcı", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Search frame
        search_frame = ttk.LabelFrame(main_frame, text="Ürün Arama", padding="10")
        search_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        search_frame.columnconfigure(1, weight=1)
        
        ttk.Label(search_frame, text="Arama:").grid(row=0, column=0, padx=(0, 10))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.search_products)
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        self.search_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        self.search_entry.bind('<Button-1>', self.on_search_click)
        self.search_entry.bind('<FocusIn>', self.on_search_focus)
        self.search_entry.bind('<Key>', self.on_search_key)
        
        ttk.Label(search_frame, text="Arama Türü:").grid(row=0, column=2, padx=(0, 10))
        self.search_type = tk.StringVar(value="stockCode")
        search_combo = ttk.Combobox(search_frame, textvariable=self.search_type, 
                                   values=["stockCode", "label"], state="readonly", width=15)
        search_combo.grid(row=0, column=3)
        
        # Control frame
        control_frame = ttk.LabelFrame(main_frame, text="Fiyat Hesaplama Kontrolleri", padding="10")
        control_frame.grid(row=1, column=3, sticky=(tk.W, tk.E, tk.N), padx=(10, 0))
        
        # Exchange rate and coefficient
        ttk.Label(control_frame, text="Kur:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.exchange_entry = ttk.Entry(control_frame, textvariable=self.exchange_rate, width=10)
        self.exchange_entry.grid(row=0, column=1, pady=2, padx=(5, 0))
        self.exchange_entry.bind('<Button-1>', self.on_exchange_click)
        self.exchange_entry.bind('<FocusIn>', self.on_exchange_focus)
        self.exchange_entry.bind('<Key>', self.on_exchange_key)
        
        ttk.Label(control_frame, text="Katsayı:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.coeff_entry = ttk.Entry(control_frame, textvariable=self.coefficient, width=10)
        self.coeff_entry.grid(row=1, column=1, pady=2, padx=(5, 0))
        self.coeff_entry.bind('<Button-1>', self.on_coeff_click)
        self.coeff_entry.bind('<FocusIn>', self.on_coeff_focus)
        self.coeff_entry.bind('<Key>', self.on_coeff_key)
        
        # Update button
        update_btn = ttk.Button(control_frame, text="Fiyatları Güncelle", 
                               command=self.update_prices)
        update_btn.grid(row=2, column=0, columnspan=2, pady=5)
        
        # Excel export button
        export_btn = ttk.Button(control_frame, text="Excel Çıktısı Ver", 
                               command=self.export_to_excel)
        export_btn.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Product details frame
        details_frame = ttk.LabelFrame(main_frame, text="Ürün Detayları", padding="10")
        details_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        details_frame.columnconfigure(0, weight=1)
        details_frame.rowconfigure(1, weight=1)
        
        # Product info
        self.product_info = tk.Text(details_frame, height=8, wrap=tk.WORD)
        self.product_info.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Price input frame
        price_frame = ttk.Frame(details_frame)
        price_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        price_frame.columnconfigure(1, weight=1)
        
        ttk.Label(price_frame, text="price1 ($):").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.price1_var = tk.DoubleVar()
        self.price1_entry = ttk.Entry(price_frame, textvariable=self.price1_var, width=15)
        self.price1_entry.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        self.price1_entry.bind('<KeyRelease>', self.calculate_prices)
        self.price1_entry.bind('<Button-1>', self.on_price1_click)
        self.price1_entry.bind('<FocusIn>', self.on_price1_focus)
        self.price1_entry.bind('<Key>', self.on_price1_key)
        
        # Calculated prices display
        prices_display = ttk.Frame(price_frame)
        prices_display.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.price_labels = {}
        price_names = ["price1 (KDV Hariç)", "price2 (1.10x)", "price3 (1.05x)", 
                      "price4 (1.15x)", "price5 (1.20x)"]
        
        for i, name in enumerate(price_names):
            ttk.Label(prices_display, text=f"{name}:").grid(row=i, column=0, sticky=tk.W, pady=2)
            self.price_labels[name] = ttk.Label(prices_display, text="0.00", 
                                              font=('Arial', 10, 'bold'), foreground='blue')
            self.price_labels[name].grid(row=i, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        
        # Add update button for single product
        update_single_btn = ttk.Button(prices_display, text="Bu Ürünü Güncelle", 
                                      command=self.update_single_product)
        update_single_btn.grid(row=len(price_names), column=0, columnspan=2, pady=(10, 0))
        
        # Treeview for product list
        tree_frame = ttk.LabelFrame(main_frame, text="Ürün Listesi", padding="10")
        tree_frame.grid(row=2, column=2, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), 
                       padx=(10, 0), pady=(0, 10))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Create treeview with checkboxes
        columns = ('select', 'stockCode', 'label', 'price1', 'price2', 'price3', 'price4', 'price5', 'buyingPrice', 'tax')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Define headings
        headings = {
            'select': 'Seç',
            'stockCode': 'stockCode',
            'label': 'label',
            'price1': 'price1',
            'price2': 'price2',
            'price3': 'price3',
            'price4': 'price4',
            'price5': 'price5',
            'buyingPrice': 'buyingPrice',
            'tax': 'tax'
        }
        
        for col in columns:
            self.tree.heading(col, text=headings[col])
            if col == 'select':
                self.tree.column(col, width=50, anchor=tk.CENTER)
            else:
                self.tree.column(col, width=100, anchor=tk.CENTER)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid treeview and scrollbars
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Bind selection event
        self.tree.bind('<<TreeviewSelect>>', self.on_product_select)
        self.tree.bind('<Button-1>', self.on_tree_click)
        
        # Add bulk update controls
        bulk_frame = ttk.Frame(tree_frame)
        bulk_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Select all button
        select_all_btn = ttk.Button(bulk_frame, text="Hepsini Seç", command=self.select_all_products)
        select_all_btn.grid(row=0, column=0, padx=(0, 5))
        
        # Deselect all button
        deselect_all_btn = ttk.Button(bulk_frame, text="Seçimi Kaldır", command=self.deselect_all_products)
        deselect_all_btn.grid(row=0, column=1, padx=(0, 5))
        
        # Bulk update price1 button
        bulk_update_btn = ttk.Button(bulk_frame, text="Seçili Ürünleri Güncelle", command=self.bulk_update_price1)
        bulk_update_btn.grid(row=0, column=2, padx=(5, 0))
        
        # Store selected items
        self.selected_items = set()
        
    def on_exchange_click(self, event):
        """Handle exchange rate entry click"""
        self.root.after(50, lambda: self.exchange_entry.selection_range(0, tk.END))
        self.root.after(50, lambda: self.exchange_entry.focus_set())
    
    def on_exchange_focus(self, event):
        """Handle exchange rate entry focus"""
        self.root.after(50, lambda: self.exchange_entry.selection_range(0, tk.END))
    
    def on_coeff_click(self, event):
        """Handle coefficient entry click"""
        self.root.after(50, lambda: self.coeff_entry.selection_range(0, tk.END))
        self.root.after(50, lambda: self.coeff_entry.focus_set())
    
    def on_coeff_focus(self, event):
        """Handle coefficient entry focus"""
        self.root.after(50, lambda: self.coeff_entry.selection_range(0, tk.END))
    
    def on_price1_click(self, event):
        """Handle price1 entry click"""
        self.root.after(50, lambda: self.price1_entry.selection_range(0, tk.END))
        self.root.after(50, lambda: self.price1_entry.focus_set())
    
    def on_price1_focus(self, event):
        """Handle price1 entry focus"""
        self.root.after(50, lambda: self.price1_entry.selection_range(0, tk.END))
    
    def on_search_click(self, event):
        """Handle search entry click"""
        self.root.after(50, lambda: self.search_entry.selection_range(0, tk.END))
        self.root.after(50, lambda: self.search_entry.focus_set())
    
    def on_search_focus(self, event):
        """Handle search entry focus"""
        self.root.after(50, lambda: self.search_entry.selection_range(0, tk.END))
    
    def on_search_key(self, event):
        """Handle search entry key press"""
        if self.search_entry.selection_present():
            self.search_entry.delete(0, tk.END)
    
    def on_exchange_key(self, event):
        """Handle exchange entry key press"""
        if self.exchange_entry.selection_present():
            self.exchange_entry.delete(0, tk.END)
    
    def on_coeff_key(self, event):
        """Handle coefficient entry key press"""
        if self.coeff_entry.selection_present():
            self.coeff_entry.delete(0, tk.END)
    
    def on_price1_key(self, event):
        """Handle price1 entry key press"""
        if self.price1_entry.selection_present():
            self.price1_entry.delete(0, tk.END)
    
    def force_initial_focus(self):
        """Force focus on search entry to make all entries responsive"""
        self.search_entry.focus_set()
        self.search_entry.selection_range(0, tk.END)
        
    def load_data(self):
        try:
            self.df = pd.read_excel('sturmmm.xlsx')
            self.filtered_df = self.df.copy()
            self.populate_tree()
            messagebox.showinfo("Başarılı", f"{len(self.df)} ürün yüklendi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyası yüklenirken hata oluştu: {str(e)}")
    
    def search_products(self, *args):
        if self.df is None:
            return
            
        search_term = self.search_var.get().lower()
        search_type = self.search_type.get()
        
        if not search_term:
            self.filtered_df = self.df.copy()
        else:
            if search_type == "stockCode":
                self.filtered_df = self.df[self.df['stockCode'].astype(str).str.lower().str.contains(search_term)]
            else:  # label
                self.filtered_df = self.df[self.df['label'].astype(str).str.lower().str.contains(search_term)]
        
        self.populate_tree()
    
    def populate_tree(self):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if self.filtered_df is None:
            return
            
        # Populate with filtered data
        for _, row in self.filtered_df.iterrows():
            stock_code = str(row['stockCode'])
            is_selected = "☑" if stock_code in self.selected_items else "☐"
            values = (
                is_selected,
                stock_code,
                str(row['label'])[:50] + "..." if len(str(row['label'])) > 50 else str(row['label']),
                f"{row['price1']:.2f}" if pd.notna(row['price1']) else "0.00",
                f"{row['price2']:.2f}" if pd.notna(row['price2']) else "0.00",
                f"{row['price3']:.2f}" if pd.notna(row['price3']) else "0.00",
                f"{row['price4']:.2f}" if pd.notna(row['price4']) else "0.00",
                f"{row['price5']:.2f}" if pd.notna(row['price5']) else "0.00",
                f"{row['buyingPrice']:.2f}" if pd.notna(row['buyingPrice']) else "0.00",
                f"{row['tax']:.1f}" if pd.notna(row['tax']) else "0.0"
            )
            self.tree.insert('', tk.END, values=values)
    
    def on_product_select(self, event):
        selection = self.tree.selection()
        if not selection:
            return
            
        item = self.tree.item(selection[0])
        stock_code = item['values'][1]  # stockCode is now in column 1 (after select column)
        
        # Find the product in dataframe
        product_mask = self.df['stockCode'] == stock_code
        if not product_mask.any():
            return
        product = self.df[product_mask].iloc[0]
        
        # Display product info
        info_text = f"""
stockCode: {product['stockCode']}
label: {product['label']}
brand: {product['brand']}
category: {product['category']}
subCategory: {product['subCategory']}
barcode: {product['barcode']}
stockAmount: {product['stockAmount']}
currencyAbbr: {product['currencyAbbr']}
tax: {product['tax']}%
        """
        
        self.product_info.delete(1.0, tk.END)
        self.product_info.insert(1.0, info_text.strip())
        
        # Set current price1 value
        if pd.notna(product['price1']):
            self.price1_var.set(self.safe_float(product['price1']))
            self.calculate_prices()
    
    def calculate_prices(self, event=None):
        try:
            price1 = self.price1_var.get()
            # Handle empty string or None values
            if price1 == "" or price1 is None:
                return
            # Convert to float safely
            try:
                price1 = self.safe_float(price1)
            except (ValueError, TypeError):
                return
            if price1 <= 0:
                return
                
            # Get current product's tax rate
            selection = self.tree.selection()
            if not selection:
                tax_rate = 10  # Default tax rate
            else:
                item = self.tree.item(selection[0])
                stock_code = item['values'][1]  # stockCode is now in column 1
                product_mask = self.df['stockCode'] == stock_code
                if product_mask.any():
                    product = self.df[product_mask].iloc[0]
                    tax_rate = self.safe_float(product['tax']) if pd.notna(product['tax']) else 10
                else:
                    tax_rate = 10  # Default tax rate
            
            # Calculate prices
            price1_excl_tax = price1 / (1 + tax_rate / 100)
            price2 = price1_excl_tax * 1.10  # KDV hariç price1'den hesapla
            price3 = price1_excl_tax * 1.05  # KDV hariç price1'den hesapla
            price4 = price1_excl_tax * 1.15  # KDV hariç price1'den hesapla
            price5 = price1_excl_tax * 1.20  # KDV hariç price1'den hesapla
            
            # Calculate buying price
            try:
                exchange_rate = self.safe_float(self.exchange_rate.get()) if self.exchange_rate.get() else 1.0
                coefficient = self.safe_float(self.coefficient.get()) if self.coefficient.get() else 1.0
            except (ValueError, TypeError):
                exchange_rate = 1.0
                coefficient = 1.0
            buying_price = exchange_rate * coefficient * price1_excl_tax
            
            # Update labels
            self.price_labels["price1 (KDV Hariç)"].config(text=f"{price1_excl_tax:.2f}")
            self.price_labels["price2 (1.10x)"].config(text=f"{price2:.2f}")
            self.price_labels["price3 (1.05x)"].config(text=f"{price3:.2f}")
            self.price_labels["price4 (1.15x)"].config(text=f"{price4:.2f}")
            self.price_labels["price5 (1.20x)"].config(text=f"{price5:.2f}")
            
        except Exception as e:
            print(f"Price calculation error: {e}")
    
    def update_prices(self):
        """Update all prices in the dataframe and refresh display"""
        try:
            try:
                exchange_rate = self.safe_float(self.exchange_rate.get()) if self.exchange_rate.get() else 1.0
                coefficient = self.safe_float(self.coefficient.get()) if self.coefficient.get() else 1.0
            except (ValueError, TypeError):
                exchange_rate = 1.0
                coefficient = 1.0
            
            if exchange_rate <= 0 or coefficient <= 0:
                messagebox.showerror("Hata", "Kur ve katsayı pozitif değerler olmalıdır.")
                return
            
            # Update buying prices for all products
            for idx, row in self.df.iterrows():
                if pd.notna(row['price1']) and pd.notna(row['tax']):
                    price1_excl_tax = row['price1'] / (1 + row['tax'] / 100)
                    self.df.at[idx, 'buyingPrice'] = exchange_rate * coefficient * price1_excl_tax
                    
                    # Update other prices (el ile girilen price1 ile çarp)
                    self.df.at[idx, 'price2'] = row['price1'] * 1.10
                    self.df.at[idx, 'price3'] = row['price1'] * 1.05
                    self.df.at[idx, 'price4'] = row['price1'] * 1.15
                    self.df.at[idx, 'price5'] = row['price1'] * 1.20
            
            # Refresh display
            self.search_products()
            messagebox.showinfo("Başarılı", "Tüm fiyatlar güncellendi.")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Fiyat güncelleme hatası: {str(e)}")
    
    def update_single_product(self):
        """Update prices for the currently selected product only"""
        try:
            # Check if a product is selected
            selection = self.tree.selection()
            if not selection:
                messagebox.showwarning("Uyarı", "Lütfen güncellenecek bir ürün seçin.")
                return
            
            # Get the selected product's stock code
            item = self.tree.item(selection[0])
            stock_code = item['values'][1]  # stockCode is now in column 1
            
            # Get new price1 value
            new_price1 = self.price1_var.get()
            if new_price1 == "" or new_price1 is None:
                messagebox.showerror("Hata", "Price1 değeri boş olamaz.")
                return
            # Convert to float safely
            try:
                new_price1 = self.safe_float(new_price1)
            except (ValueError, TypeError):
                messagebox.showerror("Hata", "Price1 geçerli bir sayı olmalıdır.")
                return
            if new_price1 <= 0:
                messagebox.showerror("Hata", "Price1 değeri 0'dan büyük olmalıdır.")
                return
            
            # Find the product in dataframe
            product_mask = self.df['stockCode'] == stock_code
            if not product_mask.any():
                messagebox.showerror("Hata", "Ürün bulunamadı.")
                return
            product_idx = self.df[product_mask].index[0]
            product = self.df.iloc[product_idx]
            
            # Get tax rate
            tax_rate = self.safe_float(product['tax']) if pd.notna(product['tax']) else 10
            
            # Calculate new prices
            price1_excl_tax = new_price1 / (1 + tax_rate / 100)
            price2 = price1_excl_tax * 1.10  # KDV hariç price1'den hesapla
            price3 = price1_excl_tax * 1.05  # KDV hariç price1'den hesapla
            price4 = price1_excl_tax * 1.15  # KDV hariç price1'den hesapla
            price5 = price1_excl_tax * 1.20  # KDV hariç price1'den hesapla
            
            # Calculate buying price
            try:
                exchange_rate = self.safe_float(self.exchange_rate.get()) if self.exchange_rate.get() else 1.0
                coefficient = self.safe_float(self.coefficient.get()) if self.coefficient.get() else 1.0
            except (ValueError, TypeError):
                exchange_rate = 1.0
                coefficient = 1.0
            buying_price = exchange_rate * coefficient * price1_excl_tax
            
            # Update the dataframe with KDV excluded price1
            self.df.at[product_idx, 'price1'] = price1_excl_tax  # KDV hariç price1
            self.df.at[product_idx, 'price2'] = price2
            self.df.at[product_idx, 'price3'] = price3
            self.df.at[product_idx, 'price4'] = price4
            self.df.at[product_idx, 'price5'] = price5
            self.df.at[product_idx, 'buyingPrice'] = buying_price
            
            # Update filtered_df as well to reflect changes
            if self.filtered_df is not None:
                filtered_mask = self.filtered_df['stockCode'] == stock_code
                if filtered_mask.any():
                    filtered_idx = self.filtered_df[filtered_mask].index[0]
                    self.filtered_df.at[filtered_idx, 'price1'] = price1_excl_tax  # KDV hariç price1
                    self.filtered_df.at[filtered_idx, 'price2'] = price2
                    self.filtered_df.at[filtered_idx, 'price3'] = price3
                    self.filtered_df.at[filtered_idx, 'price4'] = price4
                    self.filtered_df.at[filtered_idx, 'price5'] = price5
                    self.filtered_df.at[filtered_idx, 'buyingPrice'] = buying_price
            
            # Refresh the treeview
            self.populate_tree()
            
            # Clear selection after update
            self.tree.selection_remove(self.tree.selection())
            
            messagebox.showinfo("Başarılı", f"Ürün {stock_code} başarıyla güncellendi.")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Ürün güncelleme hatası: {str(e)}")
    
    def on_tree_click(self, event):
        """Handle treeview click for checkbox selection"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#1":  # First column (select column)
                item = self.tree.identify_row(event.y)
                if item:
                    stock_code = str(self.tree.item(item)['values'][1])  # stockCode is in column 1
                    if stock_code in self.selected_items:
                        self.selected_items.remove(stock_code)
                    else:
                        self.selected_items.add(stock_code)
                    self.populate_tree()
    
    def select_all_products(self):
        """Select all visible products"""
        for _, row in self.filtered_df.iterrows():
            self.selected_items.add(str(row['stockCode']))
        self.populate_tree()
    
    def deselect_all_products(self):
        """Deselect all products"""
        self.selected_items.clear()
        self.populate_tree()
    
    def bulk_update_price1(self):
        """Update price1 for all selected products"""
        try:
            if not self.selected_items:
                messagebox.showwarning("Uyarı", "Lütfen güncellenecek ürünleri seçin.")
                return
            
            # Get new price1 value from the input field
            new_price1 = self.price1_var.get()
            if new_price1 == "" or new_price1 is None:
                messagebox.showerror("Hata", "Price1 değeri boş olamaz.")
                return
            # Convert to float safely
            try:
                new_price1 = self.safe_float(new_price1)
            except (ValueError, TypeError):
                messagebox.showerror("Hata", "Price1 geçerli bir sayı olmalıdır.")
                return
            if new_price1 <= 0:
                messagebox.showerror("Hata", "Price1 değeri 0'dan büyük olmalıdır.")
                return
            
            updated_count = 0
            
            # Update each selected product
            for stock_code in self.selected_items:
                try:
                    # Ensure stock_code is string
                    stock_code = str(stock_code)
                    
                    # Find product in main dataframe using a safer approach
                    matching_rows = self.df[self.df['stockCode'].astype(str) == stock_code]
                    
                    if len(matching_rows) > 0:
                        # Get the first matching row
                        product_row = matching_rows.iloc[0]
                        product_idx = matching_rows.index[0]
                        
                    # Get tax rate
                    tax_rate = self.safe_float(product_row['tax']) if pd.notna(product_row['tax']) else 10
                    
                    # Calculate new prices
                    price1_excl_tax = new_price1 / (1 + tax_rate / 100)
                    price2 = price1_excl_tax * 1.10  # KDV hariç price1'den hesapla
                    price3 = price1_excl_tax * 1.05  # KDV hariç price1'den hesapla
                    price4 = price1_excl_tax * 1.15  # KDV hariç price1'den hesapla
                    price5 = price1_excl_tax * 1.20  # KDV hariç price1'den hesapla
                    
                    # Calculate buying price
                    try:
                        exchange_rate = self.safe_float(self.exchange_rate.get()) if self.exchange_rate.get() else 1.0
                        coefficient = self.safe_float(self.coefficient.get()) if self.coefficient.get() else 1.0
                    except (ValueError, TypeError):
                        exchange_rate = 1.0
                        coefficient = 1.0
                    buying_price = exchange_rate * coefficient * price1_excl_tax
                    
                    # Update the dataframe with KDV excluded price1
                    self.df.at[product_idx, 'price1'] = price1_excl_tax  # KDV hariç price1
                    self.df.at[product_idx, 'price2'] = price2
                    self.df.at[product_idx, 'price3'] = price3
                    self.df.at[product_idx, 'price4'] = price4
                    self.df.at[product_idx, 'price5'] = price5
                    self.df.at[product_idx, 'buyingPrice'] = buying_price
                    
                    updated_count += 1
                        
                except Exception as e:
                    print(f"Error updating product {stock_code}: {e}")
                    continue
            
            # Update filtered_df as well
            if self.filtered_df is not None:
                for stock_code in self.selected_items:
                    try:
                        # Ensure stock_code is string
                        stock_code = str(stock_code)
                        
                        # Find product in filtered dataframe
                        matching_rows = self.filtered_df[self.filtered_df['stockCode'].astype(str) == stock_code]
                        
                        if len(matching_rows) > 0:
                            filtered_idx = matching_rows.index[0]
                            product = matching_rows.iloc[0]
                            
                            tax_rate = self.safe_float(product['tax']) if pd.notna(product['tax']) else 10
                            price1_excl_tax = new_price1 / (1 + tax_rate / 100)
                            price2 = price1_excl_tax * 1.10  # KDV hariç price1'den hesapla
                            price3 = price1_excl_tax * 1.05  # KDV hariç price1'den hesapla
                            price4 = price1_excl_tax * 1.15  # KDV hariç price1'den hesapla
                            price5 = price1_excl_tax * 1.20  # KDV hariç price1'den hesapla
                            exchange_rate = self.exchange_rate.get()
                            coefficient = self.coefficient.get()
                            buying_price = exchange_rate * coefficient * price1_excl_tax
                            
                            self.filtered_df.at[filtered_idx, 'price1'] = price1_excl_tax  # KDV hariç price1
                            self.filtered_df.at[filtered_idx, 'price2'] = price2
                            self.filtered_df.at[filtered_idx, 'price3'] = price3
                            self.filtered_df.at[filtered_idx, 'price4'] = price4
                            self.filtered_df.at[filtered_idx, 'price5'] = price5
                            self.filtered_df.at[filtered_idx, 'buyingPrice'] = buying_price
                            
                    except Exception as e:
                        print(f"Error updating filtered product {stock_code}: {e}")
                        continue
            
            # Refresh the treeview
            self.populate_tree()
            
            # Clear all selections after bulk update
            self.selected_items.clear()
            
            messagebox.showinfo("Başarılı", f"{updated_count} ürün başarıyla güncellendi.")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Toplu güncelleme hatası: {str(e)}")
    
    def export_to_excel(self):
        """Export current data to Excel file"""
        try:
            if self.df is None:
                messagebox.showerror("Hata", "Önce veri yükleyin.")
                return
            
            # Ask user for file location
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Excel Dosyasını Kaydet"
            )
            
            if not filename:
                return  # User cancelled
            
            # Create a copy of the dataframe for export
            export_df = self.df.copy()
            
            # Export to Excel
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # Main data sheet
                export_df.to_excel(writer, sheet_name='Urun_Fiyatlari', index=False)
                
                # Summary sheet
                summary_data = {
                    'Toplam_Urun_Sayisi': [len(export_df)],
                    'Export_Tarihi': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                    'Kur': [self.safe_float(self.exchange_rate.get()) if self.exchange_rate.get() else 1.0],
                    'Katsayi': [self.safe_float(self.coefficient.get()) if self.coefficient.get() else 1.0],
                    'Ortalama_Price1': [export_df['price1'].mean() if not export_df['price1'].isna().all() else 0],
                    'Ortalama_Buying_Price': [export_df['buyingPrice'].mean() if not export_df['buyingPrice'].isna().all() else 0]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Ozet', index=False)
            
            messagebox.showinfo("Başarılı", f"Excel dosyası başarıyla kaydedildi:\n{filename}")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Excel çıktısı hatası: {str(e)}")

def main():
    root = tk.Tk()
    app = ProductPriceCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
