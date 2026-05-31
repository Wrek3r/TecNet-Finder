# TecNet-Finder

TecNet Finder es un analizador léxico-sintáctico desarrollado en Python para la materia de Lenguajes y Autómatas I.

El proyecto utiliza un lenguaje específico de dominio, o DSL, para escribir instrucciones relacionadas con redes IP. A partir de esas instrucciones, el sistema reconoce tokens, valida la estructura sintáctica, genera una tabla VLSM y muestra los resultados mediante una interfaz gráfica.

## Funciones principales

* Análisis léxico mediante expresiones regulares.
* Análisis sintáctico mediante parser manual descendente.
* Manejo de errores léxicos y sintácticos.
* Cálculo de subredes VLSM.
* Interfaz gráfica desarrollada con tkinter.
* Tabla de tokens reconocidos.
* Tabla de palabras reservadas.
* Árbol sintáctico visual.
* Exportación de resultados VLSM a Excel.
* Texto de ejemplo dentro del área de entrada.

## Mejoras complementarias

* Guardado del árbol sintáctico como imagen PNG.
* Botón para pegar instrucciones desde el portapapeles.
* Botones para copiar el análisis y los errores.
* Tooltips en botones pequeños.
* Organización de varias redes en hojas distintas dentro del mismo archivo Excel.

## Sintaxis del lenguaje

La estructura general de una instrucción válida es:

```txt
IP <IP_ADDRESS> MASK <SUBNET_MASK> HOSTS <NUMBER>[,<NUMBER>...] [NAME <IDENTIFIER>]
```

Ejemplo con nombre de red:

```txt
IP 192.168.1.0 MASK /24 HOSTS 50,30,10 NAME Oficina
```

Ejemplo sin nombre de red:

```txt
IP 10.0.0.0 MASK /8 HOSTS 100,200
```

## Tokens principales

| Token       | Descripción                                              |
| ----------- | -------------------------------------------------------- |
| IP          | Palabra reservada que inicia una instrucción de red.     |
| IP_ADDRESS  | Dirección IPv4 completa.                                 |
| MASK        | Palabra reservada para indicar la máscara.               |
| SUBNET_MASK | Máscara en formato CIDR, por ejemplo `/24`.              |
| HOSTS       | Palabra reservada para indicar la lista de hosts.        |
| NUMBER      | Cantidad de hosts solicitados.                           |
| COMMA       | Separador de valores en la lista de hosts.               |
| NAME        | Palabra reservada opcional para asignar nombre a la red. |
| IDENTIFIER  | Nombre asignado a la red.                                |
| DOT         | Punto usado para apoyar la detección de IP incompletas.  |

## Requisitos

Para ejecutar TecNet Finder se necesita tener instalado:

* Python 3.10 o superior.
* Dependencias indicadas en `requirements.txt`.

Las principales dependencias externas son:

* `openpyxl`
* `Pillow`

## Instalación de dependencias

Abrir una terminal dentro de la carpeta del proyecto y ejecutar:

```bash
py -m pip install -r requirements.txt
```

Si el comando anterior no funciona, usar:

```bash
python -m pip install -r requirements.txt
```

## Ejecución

Para ejecutar el compilador, usar:

```bash
py Compilador.py
```

O también:

```bash
python Compilador.py
```

## Ejemplos de entrada válida

```txt
IP 192.168.1.0 MASK /24 HOSTS 50,30,10 NAME Oficina
```

```txt
IP 10.0.0.0 MASK /8 HOSTS 100,200
```

```txt
IP 172.16.0.0 MASK /16 HOSTS 25 NAME Laboratorio
```

## Ejemplos de entrada inválida

IP incompleta:

```txt
IP 192.168.1.MASK /24 HOSTS 50
```

Orden incorrecto:

```txt
IP 192.168.1.0 HOSTS 50 MASK /24
```

Máscara mal escrita:

```txt
IP 192.168.1.0 MASK 24 HOSTS 50
```

Lista de hosts incorrecta:

```txt
IP 192.168.1.0 MASK /24 HOSTS 50,,30
```

Símbolo no reconocido:

```txt
IP 192.168.1.0 MASK /24 HOSTS 50 @
```

## Estructura recomendada del repositorio

```txt
TecNet-Finder/
│
├── Compilador.py
├── requirements.txt
├── README.md
│
└── examples/
    ├── entrada_valida.txt
    └── entradas_invalidas.txt
```

## Salidas generadas

TecNet Finder puede generar:

* Análisis léxico-sintáctico en la interfaz.
* Registro de errores léxicos y sintácticos.
* Tabla VLSM.
* Tabla de tokens.
* Tabla de palabras reservadas.
* Árbol sintáctico visual.
* Archivo Excel con los resultados VLSM.
* Imagen PNG del árbol sintáctico.

## Integrantes

* ROBLES LOERA JESÚS ANTONIO
* RODRÍGUEZ MAGAÑA ROBERTO CARLOS
* RAMÍREZ CORTEZ RICARDO SAID
* ALVARADO JIMÉNEZ EDGAR GAEL

## Materia

Lenguajes y Autómatas I
Unidad 5 - Análisis Sintáctico
Instituto Tecnológico de Tepic
