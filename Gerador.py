import os
import re
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Spinbox
from datetime import datetime
from pathlib import Path
import json
import pandas as pd
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code128
from barcode import Code128 as BC128
from barcode.writer import ImageWriter
from io import BytesIO
import math
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

class EtiquetaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de Etiquetas")
        self.resizable(False, False)
        ttk.Style(self).theme_use('clam')

        # Não registra fonte EurostileBoldExtended, pois não será usada

        # Configurações básicas
        self.output_dir    = str(Path.home() / "Documents")
        self.brand_text    = "MAHLE"
        self.subbrand_text = "MADE IN BRAZIL"

        # Variáveis de controle
        self.total_var      = tk.IntVar(value=15)
        self.use_groups_var = tk.BooleanVar(value=False)
        self.group1_count   = tk.IntVar(value=8)

        # Dados dos grupos
        for i in (1, 2):
            setattr(self, f'header{i}_var', tk.StringVar(value=""))
            setattr(self, f'piece{i}_var',  tk.StringVar(value="A 960 505 49 55"))
            setattr(self, f'date{i}_var',   tk.StringVar(value=datetime.now().strftime('%d/%m/%Y')))
            setattr(self, f'time{i}_var',   tk.StringVar(value=datetime.now().strftime('%H:%M:%S')))
            setattr(self, f'code{i}_var',   tk.StringVar(value="US873001" if i == 1 else "GH123456"))

        # Escala de preview
        self.px_mm = 4
        self.W_px  = int(74 * self.px_mm)
        self.H_px  = int(34 * self.px_mm)

        # Carrega clients.json com logos
        try:
            with open('clients.json', 'r', encoding='utf-8') as f:
                self.clients_map = json.load(f)
        except:
            self.clients_map = {}
        self.logo_paths = {1: None, 2: None}

        # Monta interface
        self._build_ui()
        self._bind_events()
        self._draw_previews()
        self._update_group_spin()

    def _build_ui(self):
        frm = ttk.LabelFrame(self, text="Configurações")
        frm.grid(row=0, column=0, padx=10, pady=10, sticky='nw')
        prw = ttk.LabelFrame(self, text="Pré-visualização")
        prw.grid(row=0, column=1, padx=10, pady=10)
        # Mensagem de aviso acima do canvas
        lbl_preview_warn = ttk.Label(prw, text="Apenas uma simples pré-visualização, terá pequenas alterações quando gerada no PDF!", foreground="#a67c00", font=("Helvetica", 9, "italic"))
        lbl_preview_warn.grid(row=0, column=0, columnspan=2, pady=(0,5))

        # Controles gerais
        ttk.Label(frm, text="Total etiquetas:").grid(row=0, column=0, sticky='e')
        Spinbox(frm, from_=1, to=100, textvariable=self.total_var, width=5).grid(row=0, column=1)
        ttk.Checkbutton(frm, text="Usar 2 grupos", variable=self.use_groups_var,
                        command=self._toggle_groups).grid(row=1, column=0, columnspan=2)
        ttk.Label(frm, text="Qtd grupo 1:").grid(row=2, column=0, sticky='e')
        self.spin1 = Spinbox(frm, from_=1, to=99, textvariable=self.group1_count, width=5)
        self.spin1.grid(row=2, column=1)

        # Grupos
        self.g1 = ttk.LabelFrame(frm, text="Grupo 1")
        self.g1.grid(row=3, column=0, columnspan=2, pady=5, sticky='ew')
        self._entry_group(self.g1, 1)
        self.g2 = ttk.LabelFrame(frm, text="Grupo 2")
        self.g2.grid(row=4, column=0, columnspan=2, pady=5, sticky='ew')
        self._entry_group(self.g2, 2)

        # Pasta e botões
        ttk.Button(frm, text="Salvar em...", command=self._choose_folder).grid(row=5, column=0, pady=5)
        self.lbl_folder = ttk.Label(frm, text=self.output_dir, width=30, anchor='w')
        self.lbl_folder.grid(row=5, column=1, pady=5)
        bf = ttk.Frame(frm)
        bf.grid(row=6, column=0, columnspan=2, pady=10)
        ttk.Button(bf, text="Gerar PDF", command=self.on_generate).pack(side='left', padx=5)
        ttk.Button(bf, text="Sair", command=self.destroy).pack(side='left')

        # Canvas de preview
        self.canvas1 = tk.Canvas(prw, width=self.W_px, height=self.H_px, bg='white')
        self.canvas1.grid(row=1, column=0, padx=5)
        self.canvas2 = tk.Canvas(prw, width=self.W_px, height=self.H_px, bg='white')
        self.canvas2.grid(row=1, column=1, padx=5)

        self._toggle_groups()

    def _entry_group(self, parent, grp):
        fields = [
            ('Cliente', f'header{grp}_var'),
            ('Peça',    f'piece{grp}_var'),
            ('Data',    f'date{grp}_var'),
            ('Hora',    f'time{grp}_var'),
            ('Código',  f'code{grp}_var')
        ]
        for i, (lbl, varname) in enumerate(fields):
            var = getattr(self, varname)
            ttk.Label(parent, text=f"{lbl}:").grid(row=i, column=0, sticky='e', padx=3)
            if lbl == 'Cliente':
                cb = ttk.Combobox(parent, textvariable=var,
                                  values=list(self.clients_map.keys()), state='readonly', width=15)
                cb.grid(row=i, column=1, padx=3)
                cb.bind('<<ComboboxSelected>>', lambda e, g=grp: self._on_client_select(g))
            else:
                ent = ttk.Entry(parent, textvariable=var, width=15)
                ent.grid(row=i, column=1, padx=3)
                if lbl == 'Data': ent.bind('<KeyRelease>', lambda e,v=var,en=ent: self._auto_format(e, v, en, 'date'))
                if lbl == 'Hora': ent.bind('<KeyRelease>', lambda e,v=var,en=ent: self._auto_format(e, v, en, 'time'))
        if grp == 1:
            ttk.Button(parent, text="Buscar Excel", command=self._lookup_to_groups).grid(
                row=len(fields), column=0, columnspan=2, pady=3)

    def _on_client_select(self, grp):
        name = getattr(self, f'header{grp}_var').get()
        self.logo_paths[grp] = self.clients_map.get(name)
        self._draw_previews()

    def _bind_events(self):
        for v in [
            self.total_var, self.use_groups_var, self.group1_count,
            self.header1_var, self.piece1_var, self.date1_var,
            self.time1_var, self.code1_var,
            self.header2_var, self.piece2_var, self.date2_var,
            self.time2_var, self.code2_var
        ]:
            v.trace_add('write', lambda *a: self._draw_previews())

    def _toggle_groups(self):
        if self.use_groups_var.get():
            self.g2.grid(); self.spin1.config(state='normal'); self.canvas2.grid()
        else:
            self.g2.grid_remove(); self.spin1.config(state='disabled'); self.canvas2.grid_remove()
        self._update_group_spin()
        self._draw_previews()

    def _choose_folder(self):
        d = filedialog.askdirectory(initialdir=self.output_dir)
        if d:
            self.output_dir = d
            self.lbl_folder.config(text=d)

    def _update_group_spin(self):
        m = max(1, self.total_var.get()-1)
        self.spin1.config(to=m)
        if self.group1_count.get() > m:
            self.group1_count.set(m)

    def _draw_previews(self):
        self._draw_canvas(self.canvas1, 1)
        if self.use_groups_var.get():
            self._draw_canvas(self.canvas2, 2)

    def _draw_canvas(self, canvas, grp):
        c = canvas
        c.delete('all')
        # Logo à esquerda, nome do cliente e subtexto à direita, todos na mesma linha no topo
        path = self.logo_paths.get(grp)
        hdr = getattr(self, f'header{grp}_var').get()
        is_daf = hdr.strip().upper().startswith('DAF')
        is_iveco = hdr.strip().upper().startswith('IVECO')
        logo_h_px = 40 if (is_daf or is_iveco) else 16  # DAF e IVECO maior, demais menor
        brand = self.brand_text
        subbrand = self.subbrand_text
        right_margin = 8
        text_x = self.W_px - right_margin
        ox_px = 2
        logo_w_px = 0
        if path and os.path.exists(path):
            orig = Image.open(path).convert('RGBA')
            # Compor sobre branco puro, eliminando qualquer pixel parcialmente transparente
            white_bg = Image.new('RGBA', orig.size, (255,255,255,255))
            white_bg.paste(orig, mask=orig.split()[3])
            # Força todos pixels não totalmente opacos para branco
            arr = white_bg.getdata()
            new_arr = [(r,g,b,255) if a==255 else (255,255,255,255) for (r,g,b,a) in arr]
            white_bg.putdata(new_arr)
            # Converte para RGB antes de redimensionar (elimina canal alfa)
            rgb_logo = white_bg.convert('RGB')
            w, h = rgb_logo.size
            new_h = logo_h_px
            new_w = int(w * new_h / h)
            img = rgb_logo.resize((new_w, new_h), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            setattr(self, f'logo_img{grp}', photo)
            c.create_image(ox_px, 5, anchor='nw', image=photo)
            logo_w_px = new_w
        # ...existing code...
        c.create_text(text_x, 7, text=brand, font=('Helvetica',12,'bold'), anchor='ne')
        c.create_text(text_x, 22, text=subbrand, font=('Helvetica',8), anchor='ne')
        # Só mostra nome do cliente se não for DAF ou IVECO
        if not (is_daf or is_iveco):
            hdr_x = ox_px + logo_w_px + 8  # 8px de espaço após logo
            c.create_text(hdr_x, 7 + logo_h_px//2, text=hdr, font=('Helvetica',12,'bold'), anchor='w')
        # Se for DAF ou IVECO, não desenha o nome do cliente, apenas a logo
        # Demais textos
        # Ajusta textos da direita para não encostarem na borda
        coords = {
            'pce': (2,12), 'shu': (7,17),
            'dat': (2,23), 'tim': (2,28)
        }
        date_val = getattr(self, f'date{grp}_var').get()
        time_val = getattr(self, f'time{grp}_var').get()
        values = {
            'pce': getattr(self, f'piece{grp}_var').get(),
            'shu': 'SHROUD',
            'dat': f'DATA: {date_val}',
            'tim': f'HORA: {time_val}'
        }
        fonts = {
            'pce':    ('Helvetica',8),
            'shu':    ('Helvetica',8),
            'dat':    ('Helvetica',7,'bold'),
            'tim':    ('Helvetica',7,'bold')
        }
        for k,(x,y) in coords.items():
            c.create_text(x*self.px_mm, y*self.px_mm, text=values[k], font=fonts[k], anchor='nw')
        # Preview barcode
        try:
            val = getattr(self, f'code{grp}_var').get() or '0000000'
            buf = BytesIO()
            bc  = BC128(str(val), writer=ImageWriter())
            bc.write(buf, {'module_width':0.2,'module_height':10,'font_size':8,'text_distance':1,'quiet_zone':1})
            buf.seek(0)
            pil = Image.open(buf).convert('RGBA')
            pil = pil.resize((int(self.W_px*0.65), pil.height), Image.LANCZOS)
            photo_bc = ImageTk.PhotoImage(pil)
            setattr(self, f'bc_img{grp}', photo_bc)
            x = self.W_px - pil.width - 2*self.px_mm
            y = int(self.H_px * 0.55)
            c.create_image(x, y, anchor='nw', image=photo_bc)
            # Legenda do código de barras sempre aparece abaixo
            legend_y = y + pil.height + 8  # espaço extra para garantir visibilidade
            c.create_text(x + pil.width/2, legend_y, text=val, font=('Helvetica',8), anchor='n')
        except Exception as e:
            print('Erro preview barcode:', e)

    def _lookup_to_groups(self):
        f = filedialog.askopenfilename(title='Excel', filetypes=[('Excel','.xlsx;.xls')])
        if not f: return
        try:
            df = pd.read_excel(f)
        except Exception as e:
            messagebox.showerror('Erro Excel', str(e))
            return
        if df.empty:
            messagebox.showinfo('Excel','vazio')
            return
        rows = [df.iloc[0]]
        if self.use_groups_var.get() and len(df) >= 2:
            rows.append(df.iloc[1])
        for idx, row in enumerate(rows, start=1):
            mapping = {
                'Cliente': f'header{idx}_var',
                'Peça':    f'piece{idx}_var',
                'Data':    f'date{idx}_var',
                'Hora':    f'time{idx}_var',
                'Código':  f'code{idx}_var'
            }
            for col, varname in mapping.items():
                v = row.get(col, '')
                if col in ('Data','Hora') and hasattr(v, 'strftime'):
                    v = v.strftime('%d/%m/%Y') if col=='Data' else v.strftime('%H:%M:%S')
                getattr(self, varname).set(str(v))
        self._draw_previews()

    def on_generate(self):
        try:
            datetime.strptime(self.date1_var.get(), '%d/%m/%Y')
            datetime.strptime(self.time1_var.get(), '%H:%M:%S')
        except:
            messagebox.showerror('Erro','Data/Hora inválidas')
            return
        if self.use_groups_var.get() and self.group1_count.get() >= self.total_var.get():
            messagebox.showerror('Erro','Configuração inválida')
            return
        self._generate_pdf()

    def _generate_pdf(self):
        total = self.total_var.get()
        g1    = self.group1_count.get() if self.use_groups_var.get() else total
        out   = os.path.join(self.output_dir, 'etiquetas.pdf')
        c     = pdf_canvas.Canvas(out, pagesize=A4)
        # Dimensões e espaçamentos
        et_w, et_h = 74*mm, 34*mm
        slot_w, slot_h = et_h, et_w
        cols, rows_per = 5, 3
        ml, mt, gh, gv = 13*mm, 12*mm, 3*mm, 3*mm
        bw, bh, maxw, rot = 0.6*mm, 10*mm, 47*mm, 90
        # Verticais
        # ...existing code...
        # Verticais
        for idx in range(total):
            grp = 1 if idx < g1 else 2
            hdr,pce,dte,tme,cds = [getattr(self, f'{nm}{grp}_var').get() for nm in ('header','piece','date','time','code')]
            is_daf = hdr.strip().upper().startswith('DAF')
            is_iveco = hdr.strip().upper().startswith('IVECO')
            logo_h_mm = 12*mm if (is_daf or is_iveco) else 6*mm
            logo_w_mm = 12*mm if (is_daf or is_iveco) else 6*mm
        for idx in range(total):
            grp = 1 if idx < g1 else 2
            hdr,pce,dte,tme,cds = [getattr(self, f'{nm}{grp}_var').get() for nm in ('header','piece','date','time','code')]
            col, row = idx % cols, idx // cols
            if row >= rows_per:
                c.showPage()
                row = 0
            x0 = ml + col*(slot_w+gh)
            y0 = A4[1] - mt - (row+1)*slot_h - row*gv
            c.saveState()
            c.translate(x0+slot_w/2, y0+slot_h/2)
            c.rotate(rot)
            c.translate(-et_w/2, -et_h/2)
            # Logo no PDF
            logo = self.logo_paths.get(grp)
            logo_drawn = False
            logo_y = et_h-5*mm-logo_h_mm
            hdr_y = et_h-5*mm
            # Verifica se é DAF (ignora espaços, variações, etc)
            is_daf = hdr.strip().upper().startswith('DAF')
            if (is_daf or is_iveco):
                hdr_x = 2*mm + logo_w_mm  # sem espaço extra
            else:
                hdr_x = 2*mm + logo_w_mm + 1*mm  # 1mm de espaço após logo
            if logo and os.path.exists(logo):
                orig = Image.open(logo).convert('RGBA')
                # Compor sobre branco puro, eliminando qualquer pixel parcialmente transparente
                white_bg = Image.new('RGBA', orig.size, (255,255,255,255))
                white_bg.paste(orig, mask=orig.split()[3])
                arr = white_bg.getdata()
                new_arr = [(r,g,b,255) if a==255 else (255,255,255,255) for (r,g,b,a) in arr]
                white_bg.putdata(new_arr)
                # Converte para RGB antes de redimensionar (elimina canal alfa)
                rgb_logo = white_bg.convert('RGB')
                # Calcula tamanho em pixels para 300 DPI
                dpi = 300
                mm_to_inch = 1/25.4
                px_w = int(logo_w_mm * mm_to_inch * dpi)
                px_h = int(logo_h_mm * mm_to_inch * dpi)
                # Redimensiona mantendo proporção, igual ao preview
                w, h = rgb_logo.size
                scale = min(px_w/w, px_h/h)
                new_h = max(1, int(h * scale))
                new_w = max(1, int(w * scale))
                logo_img = rgb_logo.resize((new_w, new_h), Image.LANCZOS)
                y_logo = hdr_y - logo_h_mm/2
                c.drawInlineImage(logo_img, 2*mm, y_logo, width=logo_w_mm, height=logo_h_mm)
            # Só desenha nome do cliente se não for DAF ou IVECO
            if not (is_daf or is_iveco):
                c.setFont('Helvetica-Bold',12)
                c.drawString(hdr_x, hdr_y, hdr)
            # Se for DAF ou IVECO, joga a peça e SHROUD mais pra esquerda
            if is_daf or is_iveco:
                pce_x = 2*mm
            else:
                pce_x = hdr_x
            c.setFont('Helvetica',8);       c.drawString(pce_x, et_h-12*mm, pce)
            c.setFont('Helvetica-Bold',8);  c.drawString(pce_x, et_h-16*mm, 'SHROUD')
            # Centraliza marca e subtexto juntos no topo direito, igual às horizontais
            c.setFont('Helvetica-Bold',9)
            brand_w = c.stringWidth(self.brand_text, 'Helvetica-Bold', 9)
            c.setFont('Helvetica',7)
            subbrand_w = c.stringWidth(self.subbrand_text, 'Helvetica', 7)
            max_w = max(brand_w, subbrand_w)
            center_x = et_w-5*mm - max_w/2
            brand_y = et_h-7*mm
            subbrand_y = brand_y - 9  # 9 pts abaixo (aprox. 2mm)
            c.setFont('Helvetica-Bold',9)
            c.drawCentredString(center_x, brand_y, self.brand_text)
            c.setFont('Helvetica',7)
            c.drawCentredString(center_x, subbrand_y, self.subbrand_text)
            c.setFont('Helvetica-Bold',7.5); c.drawString(2*mm, et_h-23*mm, f'DATA: {dte}')
            c.drawString(2*mm, et_h-28*mm, f'HORA: {tme}')
            # Barcode
            bc = code128.Code128(str(cds), barHeight=bh, barWidth=bw)
            scale = min(1.0, maxw/bc.width)
            bc_w, bc_h = bc.width*scale, bh
            bc_x, bc_y = et_w-bc_w-5*mm, 6*mm
            c.saveState()
            c.translate(bc_x, bc_y)
            c.scale(scale,1)
            bc.drawOn(c,0,0)
            c.restoreState()
            c.setFont('Helvetica-Bold',7.5); c.drawCentredString(bc_x+bc_w/2, bc_y+bc_h+2*mm, cds)
            c.restoreState()
        # Horizontais
        hor_w, hor_h = 74*mm, 34*mm
        gap = 3*mm
        # Ajusta para que as horizontais fiquem 3mm abaixo das verticais
        verticais_base_y = A4[1] - mt - rows_per*slot_h - (rows_per-1)*gv
        hor_y = verticais_base_y - hor_h - gap
        total_w = 2*hor_w + gap
        start_x = (A4[0]-total_w)/2
        for i in (0,1):
            # Se não usar 2 grupos, sempre usa grupo 1
            grp = 1 if not self.use_groups_var.get() else (1 if i==0 else 2)
            hdr,pce,dte,tme,cds = [getattr(self, f'{nm}{grp}_var').get() for nm in ('header','piece','date','time','code')]
            is_daf = hdr.strip().upper().startswith('DAF')
            is_iveco = hdr.strip().upper().startswith('IVECO')
            logo_h_mm = 12*mm if (is_daf or is_iveco) else 6*mm
            logo_w_mm = 12*mm if (is_daf or is_iveco) else 6*mm
            x = start_x + i*(hor_w+gap)
            c.saveState()
            logo = self.logo_paths.get(grp)
            ox_mm = x+2*mm
            hdr_y = hor_y+hor_h-5*mm
            # Verifica se é DAF ou IVECO
            if (is_daf or is_iveco):
                # Se for DAF ou IVECO, não desenha nome do cliente, só a logo
                hdr_x = ox_mm + logo_w_mm  # sem espaço extra
            else:
                hdr_x = ox_mm + logo_w_mm + 1*mm  # 1mm de espaço após logo
            # Sempre reserva espaço do logo
            if logo and os.path.exists(logo):
                orig = Image.open(logo).convert('RGBA')
                # Compor sobre branco puro, eliminando qualquer pixel parcialmente transparente
                white_bg = Image.new('RGBA', orig.size, (255,255,255,255))
                white_bg.paste(orig, mask=orig.split()[3])
                arr = white_bg.getdata()
                new_arr = [(r,g,b,255) if a==255 else (255,255,255,255) for (r,g,b,a) in arr]
                white_bg.putdata(new_arr)
                # Converte para RGB antes de redimensionar (elimina canal alfa)
                rgb_logo = white_bg.convert('RGB')
                # Calcula tamanho em pixels para 300 DPI
                dpi = 300
                mm_to_inch = 1/25.4
                px_w = int(logo_w_mm * mm_to_inch * dpi)
                px_h = int(logo_h_mm * mm_to_inch * dpi)
                # Redimensiona mantendo proporção
                w, h = rgb_logo.size
                scale = min(px_w/w, px_h/h)
                new_w = max(1, int(w * scale))
                new_h = max(1, int(h * scale))
                logo_img = rgb_logo.resize((new_w, new_h), Image.LANCZOS)
                y_logo = hdr_y - logo_h_mm/2
                c.drawInlineImage(logo_img, ox_mm, y_logo, width=logo_w_mm, height=logo_h_mm)
            else:
                pass  # espaço do logo já reservado por ox_mm e logo_w_mm
            # Só desenha nome do cliente se não for DAF ou IVECO
            if not (is_daf or is_iveco):
                c.setFont('Helvetica-Bold',12)
                c.drawString(hdr_x, hdr_y, hdr)
            # Centraliza marca e subtexto juntos no topo direito
            # Calcula largura dos textos
            c.setFont('Helvetica-Bold',9)
            brand_w = c.stringWidth(self.brand_text, 'Helvetica-Bold', 9)
            c.setFont('Helvetica',7)
            subbrand_w = c.stringWidth(self.subbrand_text, 'Helvetica', 7)
            max_w = max(brand_w, subbrand_w)
            # Posição X centralizada em relação ao topo direito
            center_x = x+hor_w-5*mm - max_w/2
            # Posição Y do topo
            brand_y = hor_y+hor_h-7*mm
            subbrand_y = brand_y - 9  # 9 pts abaixo (aprox. 2mm)
            c.setFont('Helvetica-Bold',9)
            c.drawCentredString(center_x, brand_y, self.brand_text)
            c.setFont('Helvetica',7)
            c.drawCentredString(center_x, subbrand_y, self.subbrand_text)
            # Número da peça e SHROUD mais próximos do topo
            # Se for DAF ou IVECO, joga a peça e SHROUD mais pra esquerda
            if is_daf or is_iveco:
                pce_x = x+2*mm
            else:
                pce_x = hdr_x
            c.setFont('Helvetica',8);      c.drawString(pce_x, hor_y+hor_h-12*mm, pce)
            c.setFont('Helvetica-Bold',8); c.drawString(pce_x, hor_y+hor_h-16*mm, 'SHROUD')
            c.setFont('Helvetica-Bold',7.5); c.drawString(x+2*mm, hor_y+hor_h-23*mm, f'DATA: {dte}')
            c.drawString(x+2*mm, hor_y+hor_h-28*mm, f'HORA: {tme}')
            # Barcode
            bc = code128.Code128(str(cds), barHeight=bh, barWidth=bw)
            scale = min(1.0, maxw/bc.width)
            bc_w, bc_h = bc.width*scale, bh
            bc_x, bc_y = x+hor_w-bc_w-5*mm, hor_y+6*mm
            c.saveState()
            c.translate(bc_x, bc_y)
            c.scale(scale,1)
            bc.drawOn(c,0,0)
            c.restoreState()
            c.setFont('Helvetica-Bold',7.5); c.drawCentredString(bc_x+bc_w/2, bc_y+bc_h+2*mm, cds)
            c.restoreState()
        c.save()
        try:
            os.startfile(out)
        except:
            subprocess.call(['xdg-open', out])
        messagebox.showinfo('Concluído', f'PDF salvo em:\n{out}')

if __name__ == '__main__':
    EtiquetaApp().mainloop()