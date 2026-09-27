# -*- coding: utf-8 -*-
from __future__ import annotations
import os, sys, traceback, subprocess, io, time
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except Exception:
    TkinterDnD=None; DND_FILES=None
from sufix_core import *

BG='#F4F7FB'; NAVY='#153F78'; BLUE='#0B57D0'; RED='#B42318'; GREEN='#217346'; GREY='#667085'; LIGHT='#E8EFF8'; WHITE='#FFFFFF'


WELCOME_VARIANTS = [
    ('welcome_background.png', 1536, 1024),
    ('welcome_background_1200.png', 1200, 800),
    ('welcome_background_1080.png', 1080, 720),
    ('welcome_background_960.png', 960, 640),
    ('welcome_background_840.png', 840, 560),
    ('welcome_background_720.png', 720, 480),
]


def choose_welcome_variant(available_width: int, available_height: int):
    """Retourne le plus grand visuel qui tient intégralement dans la zone."""
    width = max(int(available_width) - 12, 1)
    height = max(int(available_height) - 12, 1)
    for filename, image_width, image_height in WELCOME_VARIANTS:
        if image_width <= width and image_height <= height:
            return filename, image_width, image_height
    return WELCOME_VARIANTS[-1]


def welcome_control_geometry(image_x: int, image_y: int,
                             image_width: int, image_height: int):
    """Calcule les zones cliquables relativement au visuel d'accueil."""
    sx = image_width / 1536.0
    sy = image_height / 1024.0
    return {
        'version': (
            image_x + round(1236 * sx), image_y + round(22 * sy),
            round(270 * sx), round(32 * sy),
        ),
        'quit': (
            image_x + round(62 * sx), image_y + round(866 * sy),
            round(355 * sx), round(62 * sy),
        ),
        'start': (
            image_x + round(1122 * sx), image_y + round(866 * sy),
            round(350 * sx), round(62 * sy),
        ),
    }

class WizardApp:
    def __init__(self):
        self.root=(TkinterDnD.Tk() if TkinterDnD else tk.Tk())
        self.root.withdraw(); self.root.title(APP_TITLE); set_window_icon(self.root)
        self.root.update_idletasks()
        self.screen_width=max(self.root.winfo_screenwidth(), 800)
        self.screen_height=max(self.root.winfo_screenheight(), 600)
        self.compact_ui=(self.screen_height < 850 or self.screen_width < 1450)
        # Ne jamais imposer une taille minimale supérieure à la zone de travail.
        min_width=min(980, max(760, int(self.screen_width * 0.72)))
        min_height=min(650, max(520, int(self.screen_height * 0.72)))
        self.root.minsize(min_width, min_height)
        self.version=read_application_version(); self.csv_path=None; self.pdf_path=None; self.rows=[]; self.article_db={}; self.options={}; self.code_to_profile={}; self.stock_articles={}
        self.support_infos=[]; self.count_vars={}; self.spacing_vars={}; self.selected_lengths={}; self.prioritize_exact=tk.BooleanVar(value=True); self.replace_magnelis=tk.BooleanVar(value=False); self.mode_vars={'matiere':tk.BooleanVar(value=True),'equilibre':tk.BooleanVar(value=True),'fabrication':tk.BooleanVar(value=True)}
        self.output_path=None; self.generated_workbook=None; self.initial_rows=[]; self.optimized_rows=[]; self.optimized_rows_by_mode={}; self.mode_results={}; self.kpis={}; self.controls=[]
        self._configure_style(); self._show_splash(); self._show_main(); self.show_welcome()

    def _configure_style(self):
        s=ttk.Style();
        try: s.theme_use('vista')
        except: pass
        title_size=20 if self.compact_ui else 24
        h2_size=12 if self.compact_ui else 14
        body_size=10 if self.compact_ui else 11
        button_size=10 if self.compact_ui else 11
        button_padding=(14,7) if self.compact_ui else (20,10)
        s.configure('Title.TLabel',font=('Segoe UI',title_size,'bold'),foreground=NAVY,background=BG)
        s.configure('H2.TLabel',font=('Segoe UI',h2_size,'bold'),foreground=NAVY,background=BG)
        s.configure('Body.TLabel',font=('Segoe UI',body_size),background=BG)
        s.configure('Primary.TButton',font=('Segoe UI',button_size,'bold'),padding=button_padding)
        s.configure('Danger.TButton',font=('Segoe UI',button_size,'bold'),padding=button_padding,foreground=RED)

    def _show_splash(self):
        splash=tk.Toplevel(self.root); splash.overrideredirect(True); splash.attributes('-topmost',True)
        img=tk.PhotoImage(file=str(resource_path(SPLASH_FILE))); lbl=tk.Label(splash,image=img,border=0); lbl.image=img; lbl.pack()
        splash.update_idletasks(); w,h=img.width(),img.height(); x=(splash.winfo_screenwidth()-w)//2; y=(splash.winfo_screenheight()-h)//2; splash.geometry(f'{w}x{h}+{x}+{y}')
        start=time.monotonic()
        while time.monotonic()-start<4: splash.update(); time.sleep(.02)
        splash.destroy()

    def _show_main(self):
        self.root.deiconify()
        try:
            self.root.state('zoomed')
        except Exception:
            usable_height=max(self.screen_height-60, 520)
            self.root.geometry(f'{self.screen_width}x{usable_height}+0+0')
        self.root.resizable(True,True)
        self.root.update_idletasks()
        self.container=tk.Frame(self.root,bg=BG); self.container.pack(fill='both',expand=True)
        header_height=52 if self.compact_ui else 64
        footer_height=58 if self.compact_ui else 70
        self.header=tk.Frame(self.container,bg=NAVY,height=header_height); self.header.pack(fill='x'); self.header.pack_propagate(False)
        tk.Label(self.header,text='SUFIX OptimiZer',bg=NAVY,fg='white',font=('Segoe UI',16 if self.compact_ui else 19,'bold')).pack(side='left',padx=20 if self.compact_ui else 24)
        tk.Label(self.header,text=f'Version {self.version} - Août 2026',bg=NAVY,fg='white',font=('Segoe UI',9 if self.compact_ui else 10,'bold')).pack(side='right',padx=20 if self.compact_ui else 24)
        self.body=tk.Frame(self.container,bg=BG); self.body.pack(fill='both',expand=True)
        self.footer=tk.Frame(self.container,bg=WHITE,height=footer_height); self.footer.pack(fill='x'); self.footer.pack_propagate(False)
        self.status=tk.Label(self.footer,text='',bg=WHITE,fg=RED,font=('Segoe UI',10)); self.status.pack(side='left',padx=20)
        self.extra_footer_widgets=[]
        self.back_btn=ttk.Button(self.footer,text='← Retour',style='Danger.TButton'); self.next_btn=ttk.Button(self.footer,text='Suivant →',style='Primary.TButton')

    def clear(self):
        for w in self.body.winfo_children(): w.destroy()
        for w in self.extra_footer_widgets:
            try: w.destroy()
            except Exception: pass
        self.extra_footer_widgets=[]
        self.back_btn.pack_forget(); self.next_btn.pack_forget(); self.status.config(text='')

    def nav(self, back=None, nxt=None, next_text='Suivant →'):
        if back:
            self.back_btn.config(command=back); self.back_btn.pack(side='left',padx=14 if self.compact_ui else 20,pady=8 if self.compact_ui else 12)
        if nxt:
            self.next_btn.config(command=nxt,text=next_text); self.next_btn.pack(side='right',padx=14 if self.compact_ui else 20,pady=8 if self.compact_ui else 12)

    def title(self,text,sub=''):
        side_pad=26 if self.compact_ui else 42
        top_pad=16 if self.compact_ui else 30
        bottom_pad=10 if self.compact_ui else 18
        ttk.Label(self.body,text=text,style='Title.TLabel').pack(anchor='w',padx=side_pad,pady=(top_pad,4))
        if sub:
            wrap=max(620, self.root.winfo_width()-2*side_pad-40)
            ttk.Label(self.body,text=sub,style='Body.TLabel',wraplength=wrap).pack(anchor='w',padx=side_pad+2,pady=(0,bottom_pad))

    def show_welcome(self):
        self.clear()
        self.header.pack_forget()
        self.footer.pack_forget()

        canvas = tk.Canvas(self.body, bg=WHITE, highlightthickness=0)
        canvas.pack(fill='both', expand=True)
        state = {'image_id': None, 'photo': None, 'filename': None}

        version_label = tk.Label(
            canvas,
            text=f'Version {self.version} - Août 2026',
            bg='#F7F9FC',
            fg='#003B93',
            anchor='e',
        )
        quit_btn = tk.Button(
            canvas,
            text="Quitter l'application",
            command=self.root.destroy,
            bg=WHITE,
            fg='#E10600',
            activebackground='#FFF5F5',
            activeforeground='#E10600',
            relief='solid',
            bd=1,
            cursor='hand2',
        )
        start_btn = tk.Button(
            canvas,
            text='Démarrer    →',
            command=self.show_files,
            bg='#0B46B5',
            fg=WHITE,
            activebackground='#083C9D',
            activeforeground=WHITE,
            relief='flat',
            bd=0,
            cursor='hand2',
        )

        def position(event=None):
            cw=max(canvas.winfo_width(), 1)
            ch=max(canvas.winfo_height(), 1)
            filename, iw, ih = choose_welcome_variant(cw, ch)

            # Secours utile pour les anciens packages : si une vignette n'a pas
            # été embarquée, revenir au visuel principal au lieu de planter.
            image_path = resource_path(filename)
            if not image_path.exists():
                filename, iw, ih = WELCOME_VARIANTS[0]
                image_path = resource_path(filename)

            if state['filename'] != filename:
                photo=tk.PhotoImage(file=str(image_path))
                state['photo']=photo
                state['filename']=filename
                if state['image_id'] is None:
                    state['image_id']=canvas.create_image(0,0,image=photo,anchor='nw')
                else:
                    canvas.itemconfigure(state['image_id'], image=photo)

            x=max((cw-iw)//2, 0)
            y=max((ch-ih)//2, 0)
            canvas.coords(state['image_id'], x, y)

            # Les coordonnées sont exprimées proportionnellement au visuel
            # original 1536x1024 pour rester cohérentes quelle que soit la taille.
            sx=iw/1536.0
            sy=ih/1024.0
            font_scale=min(sx, sy)
            geometry = welcome_control_geometry(x, y, iw, ih)

            version_label.config(font=('Segoe UI', max(8, round(14*font_scale)), 'bold'))
            quit_btn.config(font=('Segoe UI', max(8, round(13*font_scale)), 'bold'))
            start_btn.config(font=('Segoe UI', max(9, round(14*font_scale)), 'bold'))

            for widget, key in (
                (version_label, 'version'),
                (quit_btn, 'quit'),
                (start_btn, 'start'),
            ):
                px, py, width, height = geometry[key]
                widget.place(x=px, y=py, width=width, height=height)

        canvas.bind('<Configure>', position)
        self.root.after(30, position)

    def show_files(self):
        if not self.header.winfo_ismapped():
            self.header.pack(fill='x', before=self.body)
        if not self.footer.winfo_ismapped():
            self.footer.pack(fill='x', after=self.body)
        self.clear(); self.title('1. Sélection de l’étude','Déposez le CSV ou utilisez les boutons. Le PDF associé est recherché automatiquement dans le même dossier.')
        box_height=185 if self.compact_ui else 240; side_pad=45 if self.compact_ui else 80; box=tk.Frame(self.body,bg=WHITE,highlightbackground='#9CB6D8',highlightthickness=2,height=box_height); box.pack(fill='x',padx=side_pad,pady=10 if self.compact_ui else 20); box.pack_propagate(False)
        tk.Label(box,text='Glissez-déposez ici le CSV et/ou le PDF',bg=WHITE,fg=NAVY,font=('Segoe UI',15 if self.compact_ui else 18,'bold')).pack(pady=((25 if self.compact_ui else 45),8))
        tk.Label(box,text='ou choisissez les fichiers manuellement',bg=WHITE,fg=GREY,font=('Segoe UI',11)).pack()
        buttons=tk.Frame(box,bg=WHITE); buttons.pack(pady=14 if self.compact_ui else 25)
        ttk.Button(buttons,text='Choisir le CSV',command=self.choose_csv).pack(side='left',padx=8)
        ttk.Button(buttons,text='Choisir le PDF',command=self.choose_pdf).pack(side='left',padx=8)
        if DND_FILES:
            box.drop_target_register(DND_FILES); box.dnd_bind('<<Drop>>',self.on_drop)
        self.files_label=tk.Label(self.body,text='',bg=BG,fg='#344054',font=('Segoe UI',11),justify='left'); self.files_label.pack(padx=50 if self.compact_ui else 85,anchor='w')
        self.refresh_files(); self.nav(self.show_welcome,self.prepare_inputs)

    def choose_csv(self):
        p=filedialog.askopenfilename(filetypes=[('CSV SUFIX','*.csv')]);
        if p: self.csv_path=Path(p); self.auto_pdf(); self.refresh_files()
    def choose_pdf(self):
        p=filedialog.askopenfilename(filetypes=[('PDF SUFIX','*.pdf')]);
        if p:
            self.pdf_path=Path(p); self.auto_csv(); self.refresh_files()
    def on_drop(self,event):
        paths=self.root.tk.splitlist(event.data)
        for p in paths:
            path=Path(p.strip('{}'))
            if path.suffix.lower()=='.csv': self.csv_path=path
            elif path.suffix.lower()=='.pdf': self.pdf_path=path
        if self.csv_path and not self.pdf_path: self.auto_pdf()
        if self.pdf_path and not self.csv_path: self.auto_csv()
        self.refresh_files()
    def auto_pdf(self):
        if not self.csv_path:return
        exact=self.csv_path.with_suffix('.pdf')
        if exact.exists(): self.pdf_path=exact; return
        pdfs=list(self.csv_path.parent.glob('*.pdf'))
        stem=normalize_text(self.csv_path.stem)
        ranked=sorted(((SequenceMatcher(None,stem,normalize_text(p.stem)).ratio(),p) for p in pdfs),reverse=True,key=lambda x:x[0])
        if ranked and ranked[0][0]>=.75:self.pdf_path=ranked[0][1]
    def auto_csv(self):
        if not self.pdf_path:return
        exact=self.pdf_path.with_suffix('.csv')
        if exact.exists(): self.csv_path=exact; return
        csvs=list(self.pdf_path.parent.glob('*.csv'))
        stem=normalize_text(self.pdf_path.stem)
        ranked=sorted(((SequenceMatcher(None,stem,normalize_text(p.stem)).ratio(),p) for p in csvs),reverse=True,key=lambda x:x[0])
        if ranked and ranked[0][0]>=.75:self.csv_path=ranked[0][1]

    def refresh_files(self):
        c=str(self.csv_path) if self.csv_path else 'Non sélectionné'; p=str(self.pdf_path) if self.pdf_path else 'Non détecté / facultatif si le CSV suffit'
        self.files_label.config(text=f'CSV : {c}\nPDF : {p}')

    def prepare_inputs(self):
        try:
            if not self.csv_path: raise ValueError('Sélectionnez un fichier CSV.')
            self.status.config(text='Lecture des données...'); self.root.update()
            self.rows=read_sufix_csv(self.csv_path)
            dbp=resource_path(ARTICLE_DATABASE_FILE); pp=resource_path(PROFILE_DATABASE_FILE)
            self.article_db=load_article_database(dbp); self.rows,_=apply_article_database(self.rows,self.article_db)
            self.options,self.code_to_profile=load_profile_database(pp); self.stock_articles=load_stock_article_mapping(pp)
            counts,csources,unresolved_counts=resolve_support_counts_from_csv(self.rows)
            spacings,ssources,unresolved_spacings=resolve_support_spacings_from_csv(self.rows)
            expected=ordered_unique_supports(self.rows); infos_by={}
            if self.pdf_path and self.pdf_path.exists():
                for info in extract_support_information(self.pdf_path,self.rows): infos_by[info.support]=info
            levels={str(r.get('Support') or ''):str(r.get('Niveau') or '') for r in self.rows}
            infos=[]
            for s in expected:
                pi=infos_by.get(s)
                count=counts.get(s,pi.count if pi else 1); spacing=spacings.get(s,pi.spacing_m if pi else None)
                csource=csources.get(s,pi.source if pi else 'Valeur proposée'); ssource=ssources.get(s,pi.spacing_source if pi else 'Non détecté')
                chassis=bool(pi.chassis if pi else False) and spacing is None
                infos.append(SupportInfo(s,pi.page if pi else None,count,levels.get(s,''),csource,spacing,ssource,chassis))
            self.support_infos=infos; self.show_supports()
        except Exception as exc:self.status.config(text=str(exc))

    def show_supports(self):
        self.clear(); self.title('2. Vérification des supports','Contrôlez le nombre de supports et l’espacement. Les châssis sont indiqués en gris ; leur espacement est verrouillé.')
        outer=tk.Frame(self.body,bg=BG); outer.pack(fill='both',expand=True,padx=30 if self.compact_ui else 55,pady=(0,10 if self.compact_ui else 15))
        canvas=tk.Canvas(outer,bg=BG,highlightthickness=0); sb=ttk.Scrollbar(outer,orient='vertical',command=canvas.yview); canvas.configure(yscrollcommand=sb.set); sb.pack(side='right',fill='y'); canvas.pack(side='left',fill='both',expand=True); canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        inner=tk.Frame(canvas,bg=BG); win=canvas.create_window((0,0),window=inner,anchor='nw'); inner.bind('<Configure>',lambda e:canvas.configure(scrollregion=canvas.bbox('all'))); canvas.bind('<Configure>',lambda e:canvas.itemconfigure(win,width=e.width))
        headers=['Support','Niveau','Page','Source','Nombre','Espacement (m)'];
        for j,h in enumerate(headers): tk.Label(inner,text=h,bg=NAVY,fg='white',font=('Segoe UI',10,'bold'),padx=8,pady=8).grid(row=0,column=j,sticky='nsew',padx=1,pady=1)
        self.count_vars={};self.spacing_vars={}
        for i,info in enumerate(self.support_infos,1):
            bg='#EEF1F5' if info.chassis else WHITE; fg='#98A2B3' if info.chassis else '#344054'
            vals=[info.support,info.level,'-' if info.page is None else info.page,'Châssis' if info.chassis else info.source]
            for j,v in enumerate(vals):tk.Label(inner,text=v,bg=bg,fg=fg,font=('Segoe UI',10),anchor='w',padx=8,pady=7).grid(row=i,column=j,sticky='nsew',padx=1,pady=1)
            cv=tk.StringVar(value=str(info.count)); self.count_vars[info.support]=cv; tk.Entry(inner,textvariable=cv,width=8).grid(row=i,column=4,padx=5,pady=3)
            sv=tk.StringVar(value='' if info.spacing_m is None else f'{info.spacing_m:g}'); self.spacing_vars[info.support]=sv
            ent=tk.Entry(inner,textvariable=sv,width=12,state='disabled' if info.chassis else 'normal',disabledforeground='#98A2B3'); ent.grid(row=i,column=5,padx=5,pady=3)
            if info.chassis: sv.set('Châssis')
        for j in range(6): inner.grid_columnconfigure(j,weight=1 if j in (0,1,3) else 0)
        self.nav(self.show_files,self.validate_supports)

    def validate_supports(self):
        try:
            counts={};spacings={}
            for info in self.support_infos:
                c=int(self.count_vars[info.support].get());
                if c<=0: raise ValueError(f'{info.support} : nombre de supports invalide.')
                counts[info.support]=c
                if info.chassis: spacings[info.support]=None; continue
                raw=self.spacing_vars[info.support].get().strip().replace(',','.')
                spacings[info.support]=float(raw) if raw else None
                if spacings[info.support] is not None and spacings[info.support]<=0: raise ValueError(f'{info.support} : espacement invalide.')
            self.enriched=add_quantities(self.rows,counts,spacings); self.show_optimization()
        except Exception as exc:self.status.config(text=str(exc))

    def show_optimization(self):
        self.clear()
        self.title(
            '3. Paramètres d’optimisation',
            'Choisissez les longueurs disponibles et un ou plusieurs modes de découpe. '
            'Les chemins de câbles sont automatiquement traités en barres de 3 m et '
            'n’apparaissent pas ici.'
        )
        total = build_total_for_optimization(self.enriched)
        raw = [
            row for row in total
            if is_length_item(row)
            and not is_cable_tray_code(row.get('Code article'), self.article_db)
        ]
        profiled, unmapped = attach_profiles(
            raw,
            self.options,
            self.code_to_profile,
        )
        self.unmapped = unmapped
        self.length_rows = [
            row for row in profiled
            if row.get('Catégorie profil') in {'Rail', 'Tige'}
        ]
        self.replacement = load_magnelis_replacement_mapping(
            resource_path(PROFILE_DATABASE_FILE)
        )

        top = tk.Frame(self.body, bg=BG)
        top.pack(fill='x', padx=32 if self.compact_ui else 65, pady=6 if self.compact_ui else 10)

        ttk.Checkbutton(
            top,
            text='Prioriser les longueurs exactes / prédécoupées',
            variable=self.prioritize_exact,
        ).pack(anchor='w', pady=4)

        ttk.Checkbutton(
            top,
            text='Remplacer le rail Magnelis par du rail EZ',
            variable=self.replace_magnelis,
            command=self.rebuild_length_choices,
        ).pack(anchor='w', pady=4)

        tk.Label(
            top,
            text='Modes de découpe à calculer :',
            bg=BG,
            fg=NAVY,
            font=('Segoe UI', 11, 'bold'),
        ).pack(anchor='w', pady=(12, 5))

        mode_frame = tk.Frame(top, bg=BG)
        mode_frame.pack(fill='x', anchor='w')
        for column in range(3):
            mode_frame.grid_columnconfigure(column, weight=1, uniform='modes')
        mode_cards = [
            (
                'matiere',
                'Économie matière',
                'Priorité à la longueur achetée et à la chute.',
            ),
            (
                'equilibre',
                'Équilibré',
                'Compromis matière, barres et diversité des plans.',
            ),
            (
                'fabrication',
                'Simplicité de fabrication',
                'Priorité à un nombre réduit de schémas distincts.',
            ),
        ]
        for key, label, description in mode_cards:
            card = tk.Frame(
                mode_frame,
                bg=WHITE,
                highlightbackground='#D0D5DD',
                highlightthickness=1,
                height=78 if self.compact_ui else 88,
            )
            card.grid(row=0, column=len([w for w in mode_frame.grid_slaves(row=0)]), sticky='nsew', padx=(0, 7), pady=3)
            card.grid_propagate(False)
            ttk.Checkbutton(
                card,
                text=label,
                variable=self.mode_vars[key],
            ).pack(anchor='w', padx=12, pady=(10, 3))
            tk.Label(
                card,
                text=description,
                bg=WHITE,
                fg=GREY,
                font=('Segoe UI', 9),
                justify='left',
                wraplength=210 if self.compact_ui else 250,
            ).pack(anchor='w', padx=14)

        self.length_frame = tk.Frame(self.body, bg=BG)
        self.length_frame.pack(
            fill='both',
            expand=True,
            padx=32 if self.compact_ui else 65,
            pady=4,
        )
        self.rebuild_length_choices()
        self.nav(
            self.show_supports,
            self.run_optimization,
            'Calculer les modes sélectionnés →',
        )

    def rebuild_length_choices(self):
        if not hasattr(self,'length_frame'):return
        for w in self.length_frame.winfo_children():w.destroy()
        rows=self.length_rows
        if self.replace_magnelis.get():
            preview_enriched,_=apply_magnelis_replacement(self.enriched,self.options,self.code_to_profile,self.stock_articles,self.replacement)
            preview_total=build_total_for_optimization(preview_enriched)
            preview_raw=[r for r in preview_total if is_length_item(r) and not is_cable_tray_code(r.get('Code article'),self.article_db)]
            preview_profiled,_=attach_profiles(preview_raw,self.options,self.code_to_profile)
            rows=[r for r in preview_profiled if r.get('Catégorie profil') in {'Rail','Tige'}]
        self.preview_length_rows=rows; bounds=cut_length_bounds(rows); profiles=sorted({str(r.get('_profile_id')) for r in rows if r.get('_profile_id')},key=lambda p:normalize_text(self.options[p].profile_label)); self.length_vars={}
        canvas=tk.Canvas(self.length_frame,bg=BG,highlightthickness=0);sb=ttk.Scrollbar(self.length_frame,orient='vertical',command=canvas.yview);canvas.configure(yscrollcommand=sb.set);sb.pack(side='right',fill='y');canvas.pack(side='left',fill='both',expand=True);canvas.bind_all('<MouseWheel>',lambda e:canvas.yview_scroll(int(-1*(e.delta/120)),'units'));inner=tk.Frame(canvas,bg=BG);win=canvas.create_window((0,0),window=inner,anchor='nw');inner.bind('<Configure>',lambda e:canvas.configure(scrollregion=canvas.bbox('all')));canvas.bind('<Configure>',lambda e:canvas.itemconfigure(win,width=e.width))
        for i,pid in enumerate(profiles):
            opt=self.options[pid]; mn,mx=bounds[pid]; card=tk.Frame(inner,bg=WHITE,highlightbackground='#D0D5DD',highlightthickness=1);card.pack(fill='x',pady=4)
            tk.Label(card,text=opt.profile_label,bg=WHITE,fg=NAVY,font=('Segoe UI',11,'bold')).pack(side='left',padx=12,pady=10); tk.Label(card,text=f'Découpes {mn:g} à {mx:g} m',bg=WHITE,fg=GREY).pack(side='left',padx=12)
            vars={};
            for l in opt.stock_lengths_m:
                if l+1e-9>=mn:
                    v=tk.BooleanVar(value=True);vars[float(l)]=v;ttk.Checkbutton(card,text=f'{l:g} m',variable=v).pack(side='left',padx=5)
            self.length_vars[pid]=vars

    def run_optimization(self):
        try:
            selected_modes = [
                key for key, variable in self.mode_vars.items()
                if variable.get()
            ]
            if not selected_modes:
                raise ValueError(
                    'Sélectionnez au moins un mode de découpe.'
                )

            selected = {}
            for profile_id, variables in self.length_vars.items():
                values = [
                    length
                    for length, variable in variables.items()
                    if variable.get()
                ]
                if not values:
                    raise ValueError(
                        'Sélectionnez au moins une longueur pour '
                        f'{self.options[profile_id].profile_label}.'
                    )
                selected[profile_id] = values
            self.selected_lengths = selected

            self.show_processing(
                'Calcul des modes sélectionnés et contrôles en cours...'
            )
            self.root.update()

            enriched = self.enriched
            if self.replace_magnelis.get():
                enriched, _ = apply_magnelis_replacement(
                    enriched,
                    self.options,
                    self.code_to_profile,
                    self.stock_articles,
                    self.replacement,
                )
            self.enriched = enriched

            total = build_total_for_optimization(enriched)
            cable_rows, cable_warnings = build_cable_tray_rows(
                enriched,
                self.article_db,
            )
            raw = [
                row for row in total
                if is_length_item(row)
                and not is_cable_tray_code(
                    row.get('Code article'),
                    self.article_db,
                )
            ]
            pieces = [
                row for row in total
                if not is_length_item(row)
                and not is_cable_tray_code(
                    row.get('Code article'),
                    self.article_db,
                )
            ]

            profiled, unmapped = attach_profiles(
                raw,
                self.options,
                self.code_to_profile,
            )
            length_rows = [
                row for row in profiled
                if row.get('Catégorie profil') in {'Rail', 'Tige'}
            ]
            pieces.extend(
                row for row in profiled
                if row.get('Catégorie profil') == 'Console'
            )

            # Revalider les longueurs après un éventuel remplacement Magnelis → EZ.
            bounds = cut_length_bounds(length_rows)
            selected_validated = {}
            for profile_id, (minimum, maximum) in bounds.items():
                values = self.selected_lengths.get(profile_id) or [
                    float(value)
                    for value in self.options[profile_id].stock_lengths_m
                    if value + 1e-9 >= minimum
                ]
                values = [
                    value for value in values
                    if value + 1e-9 >= minimum
                ]
                if (
                    not values
                    or max(values) + 1e-9 < maximum
                ):
                    values = [
                        float(value)
                        for value in self.options[profile_id].stock_lengths_m
                        if value + 1e-9 >= minimum
                    ]
                selected_validated[profile_id] = values

            initial = build_initial_offer_rows(
                total,
                self.article_db,
                self.code_to_profile,
                self.options,
                cable_rows,
            )
            initial_codes = sum(
                parse_float(row.get('Nombre de code à chiffrer')) or 0
                for row in initial
            )

            piece_preview = [
                {
                    'Code article': row.get('Code article'),
                    'Libellé': row.get('Libellé'),
                    'Quantité': row.get('Quantité'),
                }
                for row in pieces
                if not is_cable_tray_code(
                    row.get('Code article'),
                    self.article_db,
                )
            ]

            mode_results = {}
            mode_controls = []
            for mode_key in selected_modes:
                cut_summary, cut_detail = make_cut_results(
                    length_rows,
                    selected_validated,
                    self.options,
                    self.stock_articles,
                    self.prioritize_exact.get(),
                    mode_key,
                )
                optimized_preview = (
                    enrich_summary_rows(
                        cut_summary,
                        piece_preview,
                        self.article_db,
                    )
                    + cable_tray_offer_rows(cable_rows)
                )
                optimized_preview.sort(
                    key=lambda row: (
                        normalize_text(row.get('Category')),
                        normalize_text(row.get('Code article')),
                    )
                )
                kpis = build_optimization_kpis(
                    length_rows,
                    cut_summary,
                    cut_detail,
                )
                optimized_codes = sum(
                    parse_float(row.get('Nombre de code à chiffrer')) or 0
                    for row in optimized_preview
                )
                kpis['Codes à chiffrer initiaux'] = initial_codes
                kpis['Codes à chiffrer optimisés'] = optimized_codes
                kpis['Gain codes à chiffrer'] = (
                    initial_codes - optimized_codes
                )
                mode_results[mode_key] = {
                    'cut_summary': cut_summary,
                    'cut_detail': cut_detail,
                    'kpis': kpis,
                    'optimized_rows': optimized_preview,
                }

            versions = read_data_versions()
            base_article_version = (
                f"{article_database_version(resource_path(ARTICLE_DATABASE_FILE))}"
                f" - {versions.get('base_article_supportage', {}).get('date', '2026-08-19')}"
            )
            base_profiles_version = (
                str(
                    versions.get('sufix_profiles', {}).get(
                        'version',
                        'Package 1.1.2',
                    )
                )
                + ' - '
                + str(
                    versions.get('sufix_profiles', {}).get(
                        'date',
                        '2026-08-19',
                    )
                )
            )

            # Contrôles communs + contrôle des découpes pour chaque mode.
            first_mode = selected_modes[0]
            first_result = mode_results[first_mode]
            controls = build_control_rows(
                enriched,
                self.article_db,
                length_rows,
                unmapped,
                cable_rows,
                cable_warnings,
                first_result['cut_detail'],
                base_article_version,
                base_profiles_version,
            )
            controls = [
                row
                for row in controls
                if row.get('Contrôle') != 'Découpes affectées'
            ]
            requested = sum(
                int(round(parse_float(row.get('Quantité')) or 0))
                for row in length_rows
            )
            for mode_key in selected_modes:
                detail = mode_results[mode_key]['cut_detail']
                assigned = sum(
                    len(str(row.get('Découpes (m)') or '').split(' + '))
                    if row.get('Découpes (m)')
                    else 0
                    for row in detail
                )
                controls.append(
                    {
                        'Contrôle': (
                            'Découpes affectées - '
                            + optimization_mode_label(mode_key)
                        ),
                        'Statut': 'OK' if assigned == requested else 'ERREUR',
                        'Détail': f'{assigned}/{requested} découpes',
                    }
                )

            if cable_warnings:
                controls.insert(
                    0,
                    {
                        'Contrôle': 'Espacement chemins de câbles',
                        'Statut': 'ERREUR',
                        'Détail': ' ; '.join(cable_warnings),
                    },
                )

            wb, initial_rows, optimized_by_mode = create_output_workbook_multi(
                enriched,
                total,
                length_rows,
                pieces,
                initial,
                cable_rows,
                controls,
                mode_results,
                selected_validated,
                self.options,
                self.article_db,
                self.csv_path,
                self.pdf_path,
                self.prioritize_exact.get(),
                self.replace_magnelis.get(),
                base_article_version,
                base_profiles_version,
            )

            output = self.csv_path.with_suffix('.xlsx')
            # V1.1.3 : le classeur n'est écrit sur disque que lorsque
            # l'utilisateur clique sur « Générer l'étude excel ».
            self.generated_workbook = wb
            self.output_path = output
            self.initial_rows = initial_rows
            self.optimized_rows_by_mode = optimized_by_mode
            self.mode_results = mode_results
            self.controls = controls
            self.show_final()

        except Exception as exc:
            self.write_error(exc)
            self.show_optimization()
            self.status.config(text=str(exc), fg=RED)

    def show_processing(self,msg):
        self.clear(); self.title('Traitement',msg); pb=ttk.Progressbar(self.body,mode='indeterminate');pb.pack(fill='x',padx=50 if self.compact_ui else 100,pady=80);pb.start(10);self.root.update()

    def show_final(self):
        self.clear()
        self.title(
            'Étude terminée',
            'Comparez les modes calculés. Vous pouvez exporter l’offre initiale '
            'et/ou l’offre optimisée de chaque mode sans quitter l’application.'
        )

        status_bad = [
            row for row in self.controls
            if row.get('Statut') in {'ERREUR', 'ATTENTION'}
        ]
        banner = tk.Frame(
            self.body,
            bg='#FFF4E5' if status_bad else '#EAF7EE',
        )
        banner.pack(fill='x', padx=55, pady=(8, 6))
        tk.Label(
            banner,
            text=(
                'Contrôles : points à vérifier'
                if status_bad
                else 'Contrôles : OK'
            ),
            bg=banner['bg'],
            fg=RED if status_bad else GREEN,
            font=('Segoe UI', 13, 'bold'),
        ).pack(anchor='w', padx=18, pady=9)

        # Zone scrollable : elle reste utilisable sur des écrans plus petits.
        outer = tk.Frame(self.body, bg=BG)
        outer.pack(fill='both', expand=True, padx=45, pady=(0, 6))
        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            outer,
            orient='vertical',
            command=canvas.yview,
        )
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        inner = tk.Frame(canvas, bg=BG)
        window = canvas.create_window(
            (0, 0),
            window=inner,
            anchor='nw',
        )
        inner.bind(
            '<Configure>',
            lambda event: canvas.configure(scrollregion=canvas.bbox('all')),
        )
        canvas.bind(
            '<Configure>',
            lambda event: canvas.itemconfigure(window, width=event.width),
        )

        # En-tête comparatif.
        tk.Label(
            inner,
            text='Comparaison des modes de découpe',
            bg=BG,
            fg=NAVY,
            font=('Segoe UI', 14, 'bold'),
        ).pack(anchor='w', padx=15, pady=(4, 8))

        comparison = build_multi_mode_comparison_rows(self.mode_results)
        table = tk.Frame(
            inner,
            bg=WHITE,
            highlightbackground='#D0D5DD',
            highlightthickness=1,
        )
        table.pack(fill='x', padx=15, pady=(0, 12))

        headers = [
            'Mode',
            'Rendement',
            'Gain longueur',
            'Barres',
            'Chute',
            'Plans',
            'Points forts',
        ]
        for column, header in enumerate(headers):
            tk.Label(
                table,
                text=header,
                bg=NAVY,
                fg=WHITE,
                font=('Segoe UI', 9, 'bold'),
                padx=6,
                pady=7,
            ).grid(
                row=0,
                column=column,
                sticky='nsew',
                padx=1,
                pady=1,
            )

        for row_index, row in enumerate(comparison, start=1):
            values = [
                row.get('Mode'),
                (
                    f"{float(row.get('Rendement matière (%)') or 0):.2f}"
                    .replace('.', ',')
                    + ' %'
                ),
                f"{display_number(row.get('Gain longueur (m)'))} m",
                display_number(row.get('Barres optimisées')),
                f"{display_number(row.get('Chute optimisée (m)'))} m",
                display_number(row.get('Plans distincts')),
                row.get('Points forts') or '-',
            ]
            for column, value in enumerate(values):
                tk.Label(
                    table,
                    text=value,
                    bg=WHITE if row_index % 2 else '#F8FAFC',
                    fg=NAVY if column == 0 else '#344054',
                    font=(
                        ('Segoe UI', 9, 'bold')
                        if column == 0
                        else ('Segoe UI', 9)
                    ),
                    padx=6,
                    pady=7,
                    wraplength=210 if column == 6 else 140,
                ).grid(
                    row=row_index,
                    column=column,
                    sticky='nsew',
                    padx=1,
                    pady=1,
                )

        for column in range(len(headers)):
            table.grid_columnconfigure(
                column,
                weight=2 if column in (0, 6) else 1,
            )

        # V1.1.3 : choix unique et lisible des quatre offres. Les offres
        # optimisées ne sont proposées que si leur mode a été calculé en amont.
        tk.Label(inner, text='Choisissez les exports à générer', bg=BG, fg=NAVY,
                 font=('Segoe UI', 14, 'bold')).pack(anchor='w', padx=15, pady=(8, 4))
        tk.Label(inner, text='Cochez les offres à exporter vers EQ. Les mêmes choix déterminent les modes présents dans l’étude Excel.',
                 bg=BG, fg=GREY, font=('Segoe UI', 9)).pack(anchor='w', padx=15, pady=(0, 8))

        self.export_offer_vars = {'initial': tk.BooleanVar(value=True)}
        export_labels = [('initial', 'Offre Initiale')]
        if 'equilibre' in self.mode_results:
            self.export_offer_vars['equilibre'] = tk.BooleanVar(value=True)
            export_labels.append(('equilibre', 'Offre Equilibrée'))
        if 'fabrication' in self.mode_results:
            self.export_offer_vars['fabrication'] = tk.BooleanVar(value=True)
            export_labels.append(('fabrication', 'Offre Fabrication Simplifiée'))
        if 'matiere' in self.mode_results:
            self.export_offer_vars['matiere'] = tk.BooleanVar(value=True)
            export_labels.append(('matiere', 'Offre Economie Matière'))

        export_card = tk.Frame(inner, bg=WHITE, highlightbackground='#D0D5DD', highlightthickness=1)
        export_card.pack(fill='x', padx=15, pady=(0, 10))
        for key, label in export_labels:
            ttk.Checkbutton(export_card, text=label, variable=self.export_offer_vars[key]).pack(anchor='w', padx=18, pady=7)

        all_actions = tk.Frame(inner, bg=BG)
        all_actions.pack(fill='x', padx=15, pady=(6, 14))
        ttk.Button(all_actions, text='Générer tous les exports sélectionnés',
                   command=self.export_selected_offers, style='Primary.TButton').pack(side='left', padx=(0, 10))
        ttk.Button(all_actions, text="Générer l'étude excel",
                   command=self.generate_excel_study, style='Primary.TButton').pack(side='left')

        tk.Label(inner, text=f'Dossier de sortie : {self.output_path.parent}', bg=BG, fg='#344054',
                 font=('Segoe UI', 10)).pack(anchor='w', padx=15, pady=(0, 12))

        actions = tk.Frame(inner, bg=BG)
        actions.pack(anchor='w', padx=12, pady=(0, 24))
        ttk.Button(
            actions,
            text='Ouvrir le dossier',
            command=lambda: self.open_folder(self.output_path.parent),
        ).pack(side='left', padx=5)
        ttk.Button(
            actions,
            text='Traiter un autre dossier',
            command=self.reset_study,
            style='Primary.TButton',
        ).pack(side='left', padx=5)
        ttk.Button(
            actions,
            text="Fermer l'application",
            command=self.root.destroy,
            style='Danger.TButton',
        ).pack(side='left', padx=5)

    def _export_paths_for_mode(self, mode_key):
        date = datetime.now().strftime('%Y-%m-%d')
        stem = self.csv_path.stem
        file_suffix = optimization_mode_file_suffix(mode_key)
        return {
            'initial': (
                self.csv_path.parent
                / f'{stem}-EQ_IMPORT_OFFRE_{date}.csv'
            ),
            'optimized': (
                self.csv_path.parent
                / f'{stem}-EQ_IMPORT_OPTI_{file_suffix}_OFFRE_{date}.csv'
            ),
        }

    def export_selected_offers(self):
        try:
            created=[]
            if self.export_offer_vars.get('initial') and self.export_offer_vars['initial'].get():
                path = self._export_paths_for_mode(next(iter(self.mode_results)))['initial']
                export_to_eq_csv(path, self.initial_rows)
                created.append(path.name)
            for mode_key in ('equilibre','fabrication','matiere'):
                var=self.export_offer_vars.get(mode_key)
                if var is not None and var.get():
                    path=self._export_paths_for_mode(mode_key)['optimized']
                    export_to_eq_csv(path, self.optimized_rows_by_mode[mode_key])
                    created.append(path.name)
            if not created:
                raise ValueError('Aucun export n’est sélectionné.')
            self.status.config(text='Exports EQ créés : ' + ', '.join(created), fg=GREEN)
        except Exception as exc:
            self.status.config(text=str(exc), fg=RED)

    def generate_excel_study(self):
        try:
            if self.generated_workbook is None:
                raise ValueError("Aucune étude Excel n'est disponible.")
            selected_modes=[key for key in ('equilibre','fabrication','matiere')
                            if self.export_offer_vars.get(key) is not None and self.export_offer_vars[key].get()]
            # Clonage en mémoire pour autoriser plusieurs générations successives
            # avec des sélections différentes sans recalculer l'étude.
            buffer=io.BytesIO(); self.generated_workbook.save(buffer); buffer.seek(0)
            wb=load_workbook(buffer)
            for mode_key in list(self.mode_results.keys()):
                if mode_key not in selected_modes:
                    suffix=optimization_mode_sheet_suffix(mode_key)
                    for prefix in ('Opti globale - ', 'Plan découpe - ', 'Opti longueurs - '):
                        name=(prefix + suffix)[:31]
                        if name in wb.sheetnames:
                            wb.remove(wb[name])
            # Masquage demandé V1.1.3.
            hidden_exact={SHEET_CONTROLS, SHEET_SETTINGS, SHEET_GLOBAL, SHEET_TOTAL, SHEET_PIECES, SHEET_CABLE_TRAY}
            for ws in wb.worksheets:
                if ws.title in hidden_exact or (ws.title.startswith('Opti') and not ws.title.startswith('Opti globale')):
                    ws.sheet_state='hidden'
                else:
                    ws.sheet_state='visible'
                    configure_print_layout(ws)
            # La synthèse ne doit pas présenter de modes retirés du fichier final.
            if SHEET_COMPARISON in wb.sheetnames:
                ws=wb[SHEET_COMPARISON]
                # Conserver la structure mais masquer les lignes des modes non sélectionnés.
                labels={optimization_mode_label(k) for k in self.mode_results if k not in selected_modes}
                for row in range(2, ws.max_row+1):
                    if str(ws.cell(row,1).value or '') in labels:
                        ws.row_dimensions[row].hidden=True
            wb.save(self.output_path)
            self.status.config(text=f'Étude Excel créée : {self.output_path.name}', fg=GREEN)
        except Exception as exc:
            self.status.config(text=str(exc), fg=RED)

    def open_folder(self,path):
        try: os.startfile(str(path))
        except Exception: subprocess.Popen(['explorer',str(path)])
    def reset_study(self):
        self.csv_path=None;self.pdf_path=None;self.rows=[];self.generated_workbook=None;self.status.config(text='');
        for variable in self.mode_vars.values(): variable.set(True)
        self.show_files()
    def write_error(self,exc):
        desktop=Path.home()/'Desktop'; d=desktop if desktop.exists() else Path.home(); (d/'SUFIX_Optimiseur_erreur.txt').write_text(traceback.format_exc(),encoding='utf-8')
    def run(self): self.root.mainloop()

def main(): WizardApp().run()
