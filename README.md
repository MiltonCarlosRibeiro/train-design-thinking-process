# Design Thinking na Prática

Treinamento in-company de **12 horas (3 encontros de 4h)** sobre Design Thinking aplicado a empresas, com a identidade visual da **HayaiDataSystems** — *Dados que potencializam decisões*.

Instrutor: **Milton Carlos Ribeiro**

## Materiais prontos

| Arquivo | Para quem | O que é |
|---|---|---|
| `DT_na_Pratica_Slides-transflow.pptx` | Instrutor | Apresentação com animações, transições e frases de impacto |
| `DT_na_Pratica_Apostila.pdf` | Participantes | Apostila com conteúdo e exercícios |
| `DT_na_Pratica_Prova.pdf` | Participantes | Prova individual (6 questões, 1 por módulo) |
| `DT_na_Pratica_Gabarito.pdf` | Instrutor | Gabarito comentado |
| `DT_na_Pratica_Roteiro_Instrutor.pdf` | Instrutor | Cronograma, dinâmicas, materiais e ficha de pontuação |

## Conteúdo do treinamento

| Encontro | Módulos |
|---|---|
| 1 — Fundamentos e Empatia | 1. Design como ferramenta de inovação · 2. Empatizar |
| 2 — Definir e Idear | 3. Definir · 4. Idear |
| 3 — Prototipar, Testar e Aplicar | 5. Prototipar e testar · 6. Estudo de caso · Desafio Final · Prova |

## Como usar os slides

- As animações rodam **sozinhas** ao entrar em cada slide: basta avançar normalmente.
- Só dois momentos pedem um **clique**:
  - **Correção da prova:** o clique revela a alternativa correta (as erradas escurecem).
  - **Abertura do estudo de caso:** o clique mostra a pergunta para a turma.
- Cada módulo tem um botão **"Assistir vídeo"** que abre um slide de vídeo **oculto** (pulado na apresentação normal), com botão "Voltar".
- As notas do apresentador trazem orientações para cada slide.

## Vídeos da apresentação

Os vídeos ficam na pasta `videos` e são **vinculados** à apresentação `-transflow`. Por isso, o PPTX continua leve.

| Botão "Assistir vídeo" | Arquivo |
|---|---|
| Módulo 1 | `videos\M1.mp4` |
| Módulo 2 | `videos\M2.mp4` |
| Módulo 3 | `videos\M3.mp4` |
| Módulo 6 (modelo de pitch) | `videos\MODELOS DE PITCH\MODELO_PITCH5MIN-01-MICRONECTAR.mp4` |

**Ao copiar a pasta para outro computador** (ou mudá-la de lugar):

1. Copie a pasta `Treinamento_HayaiDataSystems` inteira, **com a pasta `videos`**.
2. Feche a apresentação, se estiver aberta.
3. Dê dois cliques em **`Configurar_Videos.bat`**. Ele aponta os vídeos de todas as apresentações desta pasta para a pasta `videos` do computador atual e mostra `ok` ou `FALTANDO` para cada vídeo.

É preciso rodar uma vez em cada computador, ou sempre que a pasta mudar de lugar. O PowerPoint só toca vídeo vinculado pelo caminho completo, e esse caminho muda de um computador para outro.

## Como gerar os arquivos

Todo o material é gerado por código a partir de `build/content.py`. Para alterar textos, edite esse arquivo e gere de novo.

```bash
pip install -r requirements.txt
python build/build_all.py
```

> **Atenção:** a geração cria `DT_na_Pratica_Slides.pptx` e **sobrescreve os PDFs**. Ajustes feitos à mão no PowerPoint (como os da versão `-transflow`) não são preservados. Feche os arquivos antes de gerar.

### Requisitos

- Python 3.10+
- Windows: os PDFs usam a fonte **Segoe UI Symbol** do sistema (marcadores ✓ ◯ ●), e os slides usam **Segoe UI**.
- A fonte dos PDFs (**Ubuntu**) já vem em `build/assets/fonts` (licença em `UFL.txt`).

## Estrutura

```
├── build/
│   ├── build_all.py       # gera slides + PDFs
│   ├── build_slides.py    # apresentação (.pptx)
│   ├── build_pdfs.py      # apostila, prova, gabarito e roteiro
│   ├── anim.py            # animações e transições dos slides
│   ├── content.py         # todo o conteúdo do treinamento
│   ├── brand.py           # cores e identidade visual
│   ├── make_assets.py     # recorta o logo
│   ├── render_pptx.ps1    # exporta os slides como PNG (requer PowerPoint)
│   └── assets/            # logos e fontes
└── *.pdf / *.pptx         # materiais prontos
```
