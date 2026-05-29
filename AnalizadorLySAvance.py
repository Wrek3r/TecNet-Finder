import re
import math
import tkinter as tk
import ipaddress
from tkinter import ttk, font, scrolledtext, messagebox

# Descripciones breves para mostrar errores más claros.
TOKEN_DESCRIPTIONS = {
    'IP': 'palabra reservada IP',
    'MASK': 'palabra reservada MASK',
    'HOSTS': 'palabra reservada HOSTS',
    'NAME': 'palabra reservada NAME',
    'IP_ADDRESS': 'dirección IP completa',
    'SUBNET_MASK': 'máscara CIDR',
    'NUMBER': 'número entero',
    'IDENTIFIER': 'identificador o nombre de red',
    'COMMA': 'coma separadora',
    'DOT': 'punto',
}

# Devuelve una descripción del token esperado.
# Para IP_ADDRESS y SUBNET_MASK intenta usar como ejemplo lo que el usuario escribió.
def describe_expected_token(token_type, current_token=None, tokens=None, pos=0):
    if token_type == 'SUBNET_MASK' and current_token and current_token[0] == 'NUMBER':
        return f"máscara CIDR, por ejemplo /{current_token[1]}"

    if token_type == 'IP_ADDRESS' and tokens is not None:
        parts = []
        i = pos

        while i < len(tokens) and tokens[i][0] in ('NUMBER', 'DOT'):
            parts.append(tokens[i][1])
            i += 1

        partial_ip = ''.join(parts)

        if partial_ip:
            if partial_ip.endswith('.'):
                partial_ip += '0'

            while partial_ip.count('.') < 3:
                partial_ip += '.0'

            octets = partial_ip.split('.')

            while len(octets) < 4:
                octets.append('0')

            fixed_octets = []
            for octet in octets[:4]:
                if octet.isdigit() and 0 <= int(octet) <= 255:
                    fixed_octets.append(octet)
                else:
                    fixed_octets.append('0')

            return f"dirección IP completa, por ejemplo {'.'.join(fixed_octets)}"

        return "dirección IP completa, por ejemplo 192.168.1.0"

    return TOKEN_DESCRIPTIONS.get(token_type, token_type)

# --- LÉXICO ---
# Analizador léxico de TecNet Finder.
# Recorre el texto ingresado y lo convierte en tokens.
class VLSMLexer:
    def __init__(self):
        # Lista de patrones léxicos.
        # Cada elemento contiene una expresión regular y el tipo de token que genera.
        # El orden es importante porque las palabras reservadas se revisan antes que IDENTIFIER.
        self.tokens = [
            # Palabras reservadas.
            # Se usa lookahead (?=...) para reconocerlas también si vienen pegadas
            # al valor esperado, por ejemplo IP192 o MASK/24.
            (r'\bIP(?=\s|\d)', 'IP'),
            (r'\bMASK(?=\s|/)', 'MASK'),
            (r'\bHOSTS(?=\s|\d|,|$)', 'HOSTS'),
            (r'\bNAME\b', 'NAME'),

            #(r'[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+', 'IP_ADDRESS'), esta solo no validaba el formato correcto de IP, por eso se reemplazó por la siguiente expresión regular más completa

            # Primero se intenta reconocer una IP completa.
            # Si la IP está incompleta, se reconocerán sus partes como NUMBER y DOT.
            (r'\b(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]|[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]|[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]|[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]|[0-9])\b', 'IP_ADDRESS'),

            (r'/\d+', 'SUBNET_MASK'),

            # Token agregado para reconocer el punto como parte válida del lenguaje.
            # Esto permite que una IP incompleta se mande al análisis sintáctico.
            (r'\.', 'DOT'),

            (r'\d+', 'NUMBER'),
            (r'[A-Za-z_][A-Za-z0-9_]*', 'IDENTIFIER'),
            (r',', 'COMMA'),
            (r'\s+', None),
        ]

    # Convierte el texto de entrada en tokens.
    # Si una parte no coincide con ningún patrón, se registra como error léxico.
    def tokenize(self, code):
        tokens = []
        errores = []
        line_num = 1
        col_num = 0
        last_type = None

        while code:
            match = None
            for pattern, token_type in self.tokens:
                regex = re.compile(pattern)
                match = regex.match(code)
                if match:
                    text = match.group(0)
                    lines = text.split("\n")
                    if len(lines) > 1:
                        line_num += len(lines) - 1
                        col_num = len(lines[-1])
                    else:
                        col_num += len(text)
                    code = code[len(text):]
                    if token_type:
                        start_col = col_num - len(text)
                        if token_type == 'IDENTIFIER':
                            if last_type == 'NAME':
                                tokens.append((token_type, text, line_num, start_col))
                                last_type = token_type
                            else:
                                errores.append(
                                    f"Identificador fuera de lugar en la línea {line_num}, posición {start_col}: '{text}'. "
                                    f"Los identificadores solo son válidos después de NAME."
                                )
                        else:
                            tokens.append((token_type, text, line_num, start_col))
                            last_type = token_type
                    break
            if not match:
                error_fragment = code.split()[0] if code.split() else code[0]
                errores.append(
                    f"Símbolo o fragmento no reconocido en la línea {line_num}, posición {col_num}: '{error_fragment}'. "
                    f"Revise que la entrada use palabras reservadas válidas como IP, MASK, HOSTS o NAME."
                )
                code = code[len(error_fragment):]
                col_num += len(error_fragment)
        return tokens, errores


# --- SINTÁCTICO ---
# Esta clase recibe los tokens generados por el lexer
# y verifica que aparezcan en el orden correcto.
class VLSMParser:
    def __init__(self, tokens):
        self.tokens = tokens
        self.pos = 0
        self.errors = []

    # Procesa todos los bloques de tokens.
    # Si un bloque es válido, lo agrega a los resultados.
    # Si hay error sintáctico, lo guarda y trata de continuar.
    def parse(self):
        results = []
        while self.pos < len(self.tokens):
            try:
                results.append(self.parse_block())
            except SyntaxError as e:
                self.errors.append(str(e))
                self.synchronize()
        return results

    # Permite recuperarse de un error sintáctico.
    # Avanza hasta encontrar otro token IP, que puede iniciar una nueva instrucción.
    def synchronize(self):
        while self.pos < len(self.tokens) and self.tokens[self.pos][0] != 'IP':
            self.pos += 1

    # Valida una instrucción completa con la estructura:
    # IP <IP_ADDRESS> MASK <SUBNET_MASK> HOSTS <lista_hosts> [NAME <IDENTIFIER>]
    def parse_block(self):
        self.expect('IP')

        # Después de IP debe venir una dirección IP completa.
        # Si aparecen NUMBER, DOT u otros tokens, el error es sintáctico,
        # porque los tokens existen, pero no forman la estructura esperada.
        if self.pos >= len(self.tokens) or self.tokens[self.pos][0] != 'IP_ADDRESS':
            token = self.tokens[self.pos] if self.pos < len(self.tokens) else (None, None, '?', '?')
            
            ejemplo_ip = describe_expected_token('IP_ADDRESS', token, self.tokens, self.pos)
            raise SyntaxError(
                f"Error sintáctico: dirección IP incompleta o mal formada. "
                f"Después de IP se esperaba por ejemplo IP_ADDRESS ({ejemplo_ip}), "
                f"pero se encontró {token[0]} ('{token[1]}') "
                f"en línea {token[2]}, posición {token[3]}."
            )

        ip_address = self.expect('IP_ADDRESS')
        self.expect('MASK')
        subnet_mask = self.expect('SUBNET_MASK')
        
        self.expect('HOSTS')
        if self.pos >= len(self.tokens) or self.tokens[self.pos][0] != 'NUMBER':
            token = self.tokens[self.pos] if self.pos < len(self.tokens) else (None, None, '?', '?')
            valor = token[1] if token[1] is not None else "fin de entrada"
            raise SyntaxError(
                f"Error sintáctico: después de HOSTS se esperaba al menos un NUMBER "
                f"con la cantidad de hosts, por ejemplo 50. "
                f"Se encontró {token[0]} ('{valor}') "
                f"en línea {token[2]}, posición {token[3]}."
            )
        num_hosts = self.parse_hosts()
        name = None

        # El nombre de la red es opcional, pero si aparece NAME,
        # obligatoriamente debe venir un IDENTIFIER después.
        if self.pos < len(self.tokens) and self.tokens[self.pos][0] == 'NAME':
            name_token = self.tokens[self.pos]
            self.pos += 1

            if self.pos >= len(self.tokens):
                raise SyntaxError(
                    f"Error sintáctico: después de NAME se esperaba IDENTIFIER "
                    f"con el nombre de la red, pero no hay más tokens "
                    f"en línea {name_token[2]}, posición {name_token[3]}."
                )

            if self.tokens[self.pos][0] != 'IDENTIFIER':
                token = self.tokens[self.pos]
                raise SyntaxError(
                    f"Error sintáctico: después de NAME se esperaba IDENTIFIER "
                    f"con el nombre de la red, pero se encontró {token[0]} ('{token[1]}') "
                    f"en línea {token[2]}, posición {token[3]}."
                )

            name = self.tokens[self.pos][1]
            self.pos += 1
        return {
            'ip_address': ip_address,
            'subnet_mask': subnet_mask,
            'num_hosts': num_hosts,
            'name': name
        }

    # Procesa la lista de hosts separados por comas.
    # Ejemplo válido: 50,30,10 se convierte en [50, 30, 10].
    def parse_hosts(self):
        hosts = []

        # El primer valor de HOSTS debe ser un número.
        token = self.tokens[self.pos]
        hosts.append(int(token[1]))
        self.pos += 1

        # Después de una coma siempre debe venir otro número.
        while self.pos < len(self.tokens) and self.tokens[self.pos][0] == 'COMMA':
            comma_token = self.tokens[self.pos]
            self.pos += 1

            if self.pos >= len(self.tokens):
                raise SyntaxError(
                    f"Error sintáctico: después de la coma se esperaba NUMBER "
                    f"con otra cantidad de hosts, pero no hay más tokens "
                    f"en línea {comma_token[2]}, posición {comma_token[3]}."
                )

            token = self.tokens[self.pos]

            if token[0] != 'NUMBER':
                raise SyntaxError(
                    f"Error sintáctico: después de la coma se esperaba NUMBER "
                    f"con otra cantidad de hosts, pero se encontró {token[0]} ('{token[1]}') "
                    f"en línea {token[2]}, posición {token[3]}."
                )

            hosts.append(int(token[1]))
            self.pos += 1

        return hosts

    # Verifica que el token actual sea del tipo esperado.
    # Si coincide, avanza; si no coincide, genera un error sintáctico.
    def expect(self, token_type):
        if self.pos < len(self.tokens):
            token = self.tokens[self.pos]

            if token[0] == token_type:
                self.pos += 1
                return token[1]
            else:
                esperado = describe_expected_token(token_type, token, self.tokens, self.pos)

                raise SyntaxError(
                    f"Error sintáctico: se esperaba {token_type} ({esperado}), "
                    f"pero se encontró {token[0]} ('{token[1]}') "
                    f"en línea {token[2]}, posición {token[3]}."
                )
        else:
            esperado = describe_expected_token(token_type)
            raise SyntaxError(
                f"Error sintáctico: se esperaba {token_type} ({esperado}), "
                f"pero no hay más tokens."
            )


# --- CÁLCULO VLSM ---
# Calcula las subredes VLSM a partir de una IP base,
# una máscara y una lista de hosts requeridos.
def calculate_vlsm(ip_address, subnet_mask, num_hosts_list, nombre_red=None):
    results = []
    current_ip = int(ipaddress.IPv4Address(ip_address))

    # Ordena los hosts de mayor a menor para asignar primero las subredes más grandes.
    sorted_hosts = sorted(num_hosts_list, reverse=True)

    for num_hosts in sorted_hosts:
        # Se suman 2 direcciones: una para red y una para broadcast.
        bits_necesarios = math.ceil(math.log2(num_hosts + 2))
        new_cidr = 32 - bits_necesarios
        block_size = 2 ** bits_necesarios

        network_address   = ipaddress.IPv4Address(current_ip)
        broadcast_address = ipaddress.IPv4Address(current_ip + block_size - 1)
        first_usable_ip   = ipaddress.IPv4Address(current_ip + 1)
        last_usable_ip    = ipaddress.IPv4Address(current_ip + block_size - 2)
        decimal_mask      = str(ipaddress.IPv4Network(f"0.0.0.0/{new_cidr}").netmask)

        # Guarda los datos calculados de la subred actual.
        results.append({
            'hosts_solicitados':      num_hosts,
            'hosts_disponibles':      block_size - 2,
            'direccion_de_red':       str(network_address),
            'nueva_mascara':          f"/{new_cidr}",
            'mascara_decimal':        decimal_mask,
            'primera_ip_utilizable':  str(first_usable_ip),
            'ultima_ip_utilizable':   str(last_usable_ip),
            'direccion_de_broadcast': str(broadcast_address),
            'ip_base':                ip_address,
            'nombre_red':             nombre_red or ip_address,
        })

        # Avanza a la siguiente dirección disponible.
        current_ip += block_size

    return results


# --- NÚMEROS DE LÍNEA ---
# Clase auxiliar para mostrar números de línea junto al área de entrada.
class TextLineNumbers(tk.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.textwidget = None

    def attach(self, text_widget):
        self.textwidget = text_widget

    def redraw(self, *args):
        self.delete("all")
        i = self.textwidget.index("@0,0")
        while True:
            dline = self.textwidget.dlineinfo(i)
            if dline is None:
                break
            y = dline[1]
            linenum = str(i).split(".")[0]
            self.create_text(2, y, anchor="nw", text=linenum, fill="#888888")
            i = self.textwidget.index(f"{i}+1line")


# --- INTERFAZ GRÁFICA ---
# Clase principal de la interfaz gráfica de TecNet Finder.
# Permite ingresar instrucciones, analizarlas y mostrar tokens, errores y tabla VLSM.
class VLSMApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TecNet Finder")
        self.root.configure(bg="#1e1e2e")

        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabelframe',       background='#1e1e2e', foreground='#cdd6f4')
        style.configure('TLabelframe.Label', background='#1e1e2e', foreground='#89b4fa', font=('Courier New', 10, 'bold'))
        style.configure('TFrame',            background='#1e1e2e')
        style.configure('TButton',           background='#313244', foreground='#cdd6f4',
                         font=('Courier New', 10), relief='flat', padding=6)
        style.map('TButton', background=[('active', '#89b4fa')], foreground=[('active', '#1e1e2e')])

        mono = font.Font(family="Courier New", size=10)

        # ── ENTRADA ──────────────────────────────────────────────
        input_frame = ttk.LabelFrame(root, text=" Entrada ")
        input_frame.pack(fill=tk.BOTH, padx=12, pady=(12, 4))

        text_frame = tk.Frame(input_frame, bg="#1e1e2e")
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.linenumbers = TextLineNumbers(text_frame, width=32, bg="#181825", bd=0, highlightthickness=0)
        self.linenumbers.pack(side="left", fill="y")

        self.input_text = scrolledtext.ScrolledText(
            text_frame, wrap=tk.WORD, width=70, height=5,
            font=mono, bg="#181825", fg="#cdd6f4",
            insertbackground="#f38ba8", selectbackground="#313244",
            relief="flat", bd=0
        )
        self.input_text.pack(side="left", fill=tk.BOTH, expand=True)
        self.input_text.bind("<KeyRelease>", self.on_key_release)
        self.linenumbers.attach(self.input_text)

        # Placeholder de ayuda
        self.hint_text = (
            "Ejemplo:\n"
            "IP 192.168.1.0 MASK /24 HOSTS 50,30,10 NAME Oficina\n"
            "IP 10.0.0.0 MASK /8 HOSTS 100,200"
        )
        self._show_hint()
        self.input_text.bind("<FocusIn>", self._clear_hint)

        # ── BOTONES ───────────────────────────────────────────────
        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=6)
        ttk.Button(btn_frame, text="▶  Analizar",       command=self.analyze).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_frame, text="🗑  Limpiar",        command=self.clear_all).pack(side=tk.LEFT, padx=6)

        # ── PESTAÑAS DE SALIDA ────────────────────────────────────
        notebook = ttk.Notebook(root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=(4, 12))
        style.configure('TNotebook',     background='#1e1e2e', borderwidth=0)
        style.configure('TNotebook.Tab', background='#313244', foreground='#cdd6f4',
                         font=('Courier New', 9), padding=[10, 4])
        style.map('TNotebook.Tab', background=[('selected', '#89b4fa')],
                  foreground=[('selected', '#1e1e2e')])

        # Pestaña 1: Tokens
        tab_tokens = ttk.Frame(notebook)
        notebook.add(tab_tokens, text='Tokens')
        self.token_text = self._make_output(tab_tokens)

        # Pestaña 2: VLSM
        tab_vlsm = ttk.Frame(notebook)
        notebook.add(tab_vlsm, text='Tabla VLSM')
        self.vlsm_text = self._make_output(tab_vlsm)

        # Pestaña 3: Errores
        tab_errors = ttk.Frame(notebook)
        notebook.add(tab_errors, text='Errores')
        self.error_text = self._make_output(tab_errors, fg="#f38ba8")
        ttk.Button(tab_errors, text="Borrar Errores", command=self.clear_errors).pack(pady=4)

        self.vlsm_data = None

    # ── HELPERS ───────────────────────────────────────────────────
    def _make_output(self, parent, fg="#cdd6f4"):
        t = scrolledtext.ScrolledText(
            parent, wrap=tk.WORD, width=80, height=12,
            font=font.Font(family="Courier New", size=10),
            bg="#181825", fg=fg,
            insertbackground="#f38ba8", state=tk.DISABLED,
            relief="flat", bd=4
        )
        t.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        return t

    def _write(self, widget, text, clear=False):
        widget.config(state=tk.NORMAL)
        if clear:
            widget.delete("1.0", tk.END)
        widget.insert(tk.END, text)
        widget.config(state=tk.DISABLED)

    # Borra el texto de ejemplo cuando el usuario hace clic en el área de entrada.
    def _clear_hint(self, event):
        if self._hint_active:
            self.input_text.delete("1.0", tk.END)
            self.input_text.config(fg="#cdd6f4")
            self._hint_active = False
            self.input_text.mark_set("insert", "1.0")
            self.root.after_idle(self.linenumbers.redraw)

    # Muestra el texto de ejemplo al iniciar el programa
    # y cuando el usuario presiona el botón Limpiar.
    def _show_hint(self):
        self.input_text.config(state=tk.NORMAL)
        self.input_text.delete("1.0", tk.END)
        self.input_text.insert("1.0", self.hint_text)
        self.input_text.config(fg="#585b70")
        self._hint_active = True
        self.root.after_idle(self.linenumbers.redraw)

    # ── EVENTOS ───────────────────────────────────────────────────
    def on_key_release(self, event=None):
        self.linenumbers.redraw()
        self.highlight_reserved_words()

    # Resalta visualmente las palabras reservadas IP, MASK, HOSTS y NAME.
    def highlight_reserved_words(self):
        colors = {'IP': '#89b4fa', 'MASK': '#a6e3a1', 'HOSTS': '#fab387', 'NAME': '#f9e2af'}
        for word, color in colors.items():
            tag = f'rw_{word}'
            self.input_text.tag_remove(tag, '1.0', tk.END)
            idx = '1.0'
            while True:
                idx = self.input_text.search(r'\b' + word + r'\b', idx, tk.END, regexp=True)
                if not idx:
                    break
                end = f"{idx}+{len(word)}c"
                self.input_text.tag_add(tag, idx, end)
                self.input_text.tag_config(tag, foreground=color, font=font.Font(family="Courier New", size=10, weight="bold"))
                idx = end

    # ── ANÁLISIS ──────────────────────────────────────────────────
    # Ejecuta el flujo principal:
    # 1. Lee la entrada del usuario.
    # 2. Realiza el análisis léxico.
    # 3. Realiza el análisis sintáctico.
    # 4. Calcula la tabla VLSM si no hay errores.
    # 5. Muestra los resultados en la interfaz.
    def analyze(self):
        if self._hint_active:
            messagebox.showwarning("Advertencia", "Introduce un código para analizar.")
            return

        code = self.input_text.get("1.0", tk.END).strip()
        if not code:
            messagebox.showwarning("Advertencia", "El área de entrada está vacía.")
            return

        self.clear_errors()
        self._write(self.token_text, "", clear=True)
        self._write(self.vlsm_text,  "", clear=True)
        self.vlsm_data = None

        # ── 1. LÉXICO ──
        lexer = VLSMLexer()
        tokens, lex_errors = lexer.tokenize(code)

        token_out = "=== ANÁLISIS LÉXICO ===\n"
        if lex_errors:
            self._write(self.error_text, "=== ERRORES LÉXICOS ===\n")
            for e in lex_errors:
                self._write(self.error_text, f"  {e}\n")
            token_out += f"  {len(lex_errors)} error(es) léxico(s) encontrado(s).\n"
        else:
            for t in tokens:
                ttype, val, line, col = t
                token_out += f"  [{ttype}]  '{val}'  →  línea {line}, col {col}\n"

        self._write(self.token_text, token_out, clear=True)

        if lex_errors:
            return  # No continuar si hay errores léxicos

        # ── 2. SINTÁCTICO ──
        parser = VLSMParser(tokens)
        blocks = parser.parse()

        if parser.errors:
            self._write(self.error_text, "\n=== ERRORES SINTÁCTICOS ===\n")
            for e in parser.errors:
                self._write(self.error_text, f"  {e}\n")

        if not blocks:
            self._write(self.vlsm_text, "No se encontraron bloques válidos.\n", clear=True)
            return

        # ── 3. CÁLCULO VLSM (usando blocks directamente, sin semántico) ──
        all_results = []
        for block in blocks:
            try:
                result = calculate_vlsm(
                    block['ip_address'],
                    block['subnet_mask'],
                    block['num_hosts'],
                    nombre_red=block.get('name')
                )
                all_results.extend(result)
            except Exception as ex:
                self._write(self.error_text, f"\n  Error al calcular VLSM para {block['ip_address']}: {ex}\n")

        self.vlsm_data = all_results

        # ── 4. MOSTRAR TABLA VLSM ──
        vlsm_out = "=== TABLA VLSM ===\n\n"
        grouped = {}
        for r in all_results:
            key = r.get('nombre_red') or r['ip_base']
            grouped.setdefault(key, []).append(r)

        for nombre, subredes in grouped.items():
            vlsm_out += f"  Red: {nombre}\n"

            header = (
                f"  {'#':<4} {'Hosts Solic.':<14} {'Hosts Disp.':<13} "
                f"{'Dirección Red':<18} {'CIDR':<8} {'Máscara':<18} "
                f"{'Primera IP':<16} {'Última IP':<16} {'Broadcast'}\n"
            )
            separator = "  " + "─" * (len(header) + 3) + "\n"
            vlsm_out += separator
            vlsm_out += header
            vlsm_out += separator

            for i, s in enumerate(subredes, 1):
                vlsm_out += (
                    f"  {i:<4} {s['hosts_solicitados']:<14} {s['hosts_disponibles']:<13} "
                    f"{s['direccion_de_red']:<18} {s['nueva_mascara']:<8} {s['mascara_decimal']:<18} "
                    f"{s['primera_ip_utilizable']:<16} {s['ultima_ip_utilizable']:<16} {s['direccion_de_broadcast']}\n"
                )
            vlsm_out += "\n"

        self._write(self.vlsm_text, vlsm_out, clear=True)

        if not parser.errors:
            messagebox.showinfo("Análisis completo", f"✅ {len(all_results)} subred(es) calculada(s) correctamente.")

    # ── LIMPIAR ───────────────────────────────────────────────────
    # Limpia el área donde se muestran los errores.
    def clear_errors(self):
        self._write(self.error_text, "", clear=True)

    # Limpia tokens, errores, tabla VLSM y restaura el texto de ejemplo.
    def clear_all(self):
        self.clear_errors()
        self._write(self.token_text, "", clear=True)
        self._write(self.vlsm_text, "", clear=True)
        self.vlsm_data = None
        self._show_hint()

# Punto de entrada del programa.
# Este bloque solo se ejecuta cuando el archivo se abre directamente,
# no cuando se importa desde otro archivo.
if __name__ == "__main__":
    root = tk.Tk() # Crea la ventana principal de la interfaz gráfica.
    root.minsize(800, 600) # Define el tamaño mínimo de la ventana.
    app = VLSMApp(root) # Crea la aplicación TecNet Finder dentro de la ventana principal.
    root.mainloop() # Mantiene abierta la ventana y espera las acciones del usuario.
