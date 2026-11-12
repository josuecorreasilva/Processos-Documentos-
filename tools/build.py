#!/usr/bin/env python3
"""Gera as páginas estáticas do site Queremos Viver — PET-Saúde Clima UCPel."""

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

ICONS = {
    "home": '<svg class="icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 11l9-8 9 8"/><path d="M5 10v10h5v-6h4v6h5V10"/></svg>',
    "info": '<svg class="icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/></svg>',
    "people": '<svg class="icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="9" cy="8" r="3.2"/><path d="M3 20c0-3.3 2.7-5.5 6-5.5s6 2.2 6 5.5"/><path d="M16 5.2a3.2 3.2 0 0 1 0 6M18 14.8c2 .6 3.5 2.3 3.5 5.2"/></svg>',
    "calendar": '<svg class="icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3.5" y="5" width="17" height="15" rx="2.5"/><path d="M3.5 10h17M8 3v4M16 3v4"/></svg>',
    "news": '<svg class="icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 4h12a2 2 0 0 1 2 2v13a1 1 0 0 1-1 1H7a2 2 0 0 1-2-2z"/><path d="M9 8h6M9 12h6M9 16h4"/></svg>',
    "docs": '<svg class="icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 3h7l5 5v13a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z"/><path d="M14 3v5h5M9 13h6M9 17h6"/></svg>',
    "arrow": '<svg class="icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12h14M13 6l6 6-6 6"/></svg>',
    "ig": '<svg class="icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3.5" y="3.5" width="17" height="17" rx="5"/><circle cx="12" cy="12" r="4"/><path d="M17.2 6.8h.01"/></svg>',
    "menu": '<svg class="icon" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 7h16M4 12h16M4 17h16"/></svg>',
    "close": '<svg class="icon" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    "chev": '<svg class="lnav__seta" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg>',
}


def depth_of(rel_path: str) -> int:
    parts = Path(rel_path).parts
    return max(0, len(parts) - 1)


def prefix(rel_path: str) -> str:
    d = depth_of(rel_path)
    return "" if d == 0 else "../" * d


def nav(rel_path: str, current: str) -> str:
    p = prefix(rel_path)
    def cur(key: str) -> str:
        return ' aria-current="page"' if current == key else ""

    projeto_open = " open data-fixo" if current.startswith("projeto") or current.startswith("eixo") or current == "documentos" else ""
    equipe_open = " open data-fixo" if current.startswith("equipe") else ""

    return f"""
    <aside class="lateral-nav" id="menu">
      <a class="lateral-nav__marca" href="{p}" aria-label="Queremos Viver — página inicial">
        <img class="marca-svg" src="{p}static/img/marca-queremos-viver.svg" width="240" height="240" alt="">
        <strong>Queremos <span>Viver</span></strong>
        <small>PET-Saúde Clima · UCPel</small>
      </a>
      <nav aria-label="Principal">
        <ul class="lnav">
          <li><a href="{p}"{cur("home")}>{ICONS["home"]}Início</a></li>
          <li>
            <details class="lnav__det"{projeto_open}>
              <summary class="lnav__grupo">{ICONS["info"]}<span>O projeto</span>{ICONS["chev"]}</summary>
              <ul class="lnav__eixos">
                <li><a class="cor-azul" href="{p}o-projeto/"{cur("projeto")}><span class="pt" aria-hidden="true"></span>Apresentação</a></li>
                <li><a class="cor-azul" href="{p}eixos/eixo-1/"{cur("eixo-1")}><span class="pt" aria-hidden="true"></span>Eixo 1 · Cuidado no território</a></li>
                <li><a class="cor-vermelho" href="{p}eixos/eixo-2/"{cur("eixo-2")}><span class="pt" aria-hidden="true"></span>Eixo 2 · Acesso e integralidade</a></li>
                <li><a class="cor-amarelo" href="{p}eixos/eixo-3/"{cur("eixo-3")}><span class="pt" aria-hidden="true"></span>Eixo 3 · Comunicação e inovação</a></li>
                <li><a class="cor-azul" href="{p}documentos/"{cur("documentos")}><span class="pt" aria-hidden="true"></span>Documentos</a></li>
              </ul>
            </details>
          </li>
          <li>
            <details class="lnav__det"{equipe_open}>
              <summary class="lnav__grupo">{ICONS["people"]}<span>Equipe</span>{ICONS["chev"]}</summary>
              <ul class="lnav__eixos">
                <li><a class="cor-verde" href="{p}equipe/"{cur("equipe")}><span class="pt" aria-hidden="true"></span>Quem somos</a></li>
                <li><a class="cor-verde" href="{p}equipe/em-numeros/"{cur("equipe-numeros")}><span class="pt" aria-hidden="true"></span>A equipe em números</a></li>
              </ul>
            </details>
          </li>
          <li><a href="{p}agenda/"{cur("agenda")}>{ICONS["calendar"]}Agenda</a></li>
          <li><a href="{p}noticias/"{cur("noticias")}>{ICONS["news"]}Notícias</a></li>
          <li class="lnav__sec"><a href="{p}sobre/"{cur("sobre")}>{ICONS["info"]}Sobre o site</a></li>
        </ul>
      </nav>
      <div class="lateral-nav__rodape">
        <a class="lateral-nav__social" href="https://www.instagram.com/ucpel/" target="_blank" rel="noopener">{ICONS["ig"]}<span>Instagram<small>@ucpel</small></span></a>
        <p>PET-Saúde: Clima<br>UCPel · Pelotas/RS</p>
      </div>
    </aside>
"""


def shell(rel_path: str, *, title: str, description: str, current: str, body_class: str, main: str) -> str:
    p = prefix(rel_path)
    return f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{title}</title>
  <meta name="description" content="{description}">
  <meta name="theme-color" content="#0E6F8F">
  <meta property="og:type" content="website">
  <meta property="og:title" content="{title.split(' · ')[0]}">
  <meta property="og:description" content="{description}">
  <meta property="og:locale" content="pt_BR">
  <link rel="icon" type="image/svg+xml" href="{p}static/img/marca-queremos-viver.svg">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Manrope:wght@400;500;600;700;800&display=swap">
  <link rel="stylesheet" href="{p}static/css/site.css">
</head>
<body{f' class="{body_class}"' if body_class else ''}>
  <a class="skip" href="#conteudo">Ir para o conteúdo</a>

  <header class="topo-mobile">
    <a class="marca marca--txt" href="{p}" aria-label="Queremos Viver — página inicial"><span class="marca-nome">Queremos <b>Viver</b></span></a>
    <button class="menu-btn" type="button" aria-expanded="false" aria-controls="menu">
      <span class="menu-btn__abrir">{ICONS["menu"]}<span class="sr">Abrir menu</span></span>
      <span class="menu-btn__fechar">{ICONS["close"]}<span class="sr">Fechar menu</span></span>
    </button>
  </header>

  <div class="app">
    {nav(rel_path, current)}
    <div class="area">
      <main id="conteudo">
{main}
      </main>
      <footer class="rodape">
        <div class="wrap rodape__linha">
          <img class="rodape__ucpel" src="{p}static/img/logo-ucpel.png" alt="Universidade Católica de Pelotas" width="203" height="113" loading="lazy">
          <p class="rodape__contato">Pelotas/RS · parceria com a SMS Pelotas · <a href="https://www.instagram.com/ucpel/" target="_blank" rel="noopener">@ucpel</a></p>
          <p class="rodape__copy">© 2026 Queremos Viver · PET-Saúde Clima UCPel</p>
        </div>
      </footer>
    </div>
  </div>
  <script src="{p}static/js/main.js" defer></script>
</body>
</html>
"""


DESC = (
    "Queremos Viver (PET-Saúde: Clima, UCPel): articulação interprofissional para adaptação "
    "e resiliência do SUS frente às emergências climáticas em Pelotas/RS."
)


def page_cab(etiqueta: str, titulo: str, texto: str, cor: str = "cor-azul") -> str:
    return f"""<section class="pagina-cab {cor}">
  <div class="wrap">
    <p class="etiqueta">{etiqueta}</p>
    <h1>{titulo}</h1>
    <p class="pagina-cab__texto">{texto}</p>
  </div>
  <svg class="pagina-cab__onda" viewBox="0 0 1440 60" preserveAspectRatio="none" aria-hidden="true" focusable="false"><path d="M0 30C260 70 520 0 780 30S1200 62 1440 24V60H0z" fill="#f4f7f6"/></svg>
</section>
"""


def barra(label: str, value: int, max_v: int) -> str:
    pct = 0 if max_v == 0 else round(100 * value / max_v)
    return f'<div class="barra"><span>{label}</span><i><em style="width:{pct}%"></em></i><b>{value}</b></div>'


PAGES = {}

PAGES["index.html"] = {
    "title": "Queremos Viver — PET-Saúde Clima UCPel",
    "current": "home",
    "body_class": "home",
    "main": f"""
<section class="inicio" aria-labelledby="t-titulo">
  <div class="inicio__texto">
    <p class="etiqueta">PET-Saúde: Clima · UCPel · Pelotas/RS</p>
    <h1 id="t-titulo">Queremos <span>Viver</span></h1>
    <p class="inicio__sub">Articulação interprofissional para adaptação e resiliência do SUS frente às emergências climáticas</p>
    <p class="inicio__txt">Ensino, serviço e comunidade juntos para proteger territórios vulneráveis de Pelotas diante de enchentes e eventos extremos.</p>
    <div class="inicio__acoes">
      <a class="btn btn--primario" href="o-projeto/">Conheça o projeto {ICONS["arrow"]}</a>
      <a class="btn btn--sec" href="equipe/em-numeros/">A equipe em números</a>
    </div>
  </div>
  <img class="inicio__marca" src="static/img/marca-queremos-viver.svg" width="240" height="240" alt="Marca Queremos Viver">
</section>
<section class="home-destaques" aria-label="Atalhos">
  <a class="home-destaque" href="eixos/eixo-1/"><strong>Eixo 1</strong><span>Cuidado no território e vigilância em saúde ambiental.</span></a>
  <a class="home-destaque" href="eixos/eixo-2/"><strong>Eixo 2</strong><span>Acesso à atenção especializada e integralidade do cuidado.</span></a>
  <a class="home-destaque" href="eixos/eixo-3/"><strong>Eixo 3</strong><span>Comunicação, letramento digital e justiça climática.</span></a>
</section>
""",
}

PAGES["o-projeto/index.html"] = {
    "title": "O projeto · Queremos Viver — PET-Saúde Clima UCPel",
    "current": "projeto",
    "body_class": "",
    "main": page_cab(
        "O projeto",
        "Apresentação",
        "O PET-Saúde Clima da UCPel integra ensino, serviço e comunidade para fortalecer a resposta do SUS em Pelotas.",
    )
    + f"""
<section class="secao secao--curta">
  <div class="wrap prosa">
    <p><strong>Queremos Viver</strong> é o projeto da Universidade Católica de Pelotas selecionado no Programa PET-Saúde: Clima (2026–2028), do Ministério da Saúde, em parceria com a Secretaria Municipal de Saúde de Pelotas.</p>
    <p>Sob coordenação da professora <strong>Luísa Jardim Corrêa de Oliveira</strong>, a proposta prioriza comunidades mais expostas a alagamentos, com mapeamento de risco, prevenção e ações interprofissionais em situações graves.</p>
    <p>As atividades se organizam em cinco Grupos de Aprendizagem Tutorial (GATs), distribuídos pelos três eixos nacionais do programa.</p>
  </div>
  <div class="wrap" style="margin-top:28px">
    <div class="grade-eixos">
      <a class="eixo-card cor-azul" href="../eixos/eixo-1/"><small>Eixo I</small><h3>Cuidado no território</h3><p>Habitat resiliente, saúde ambiental e vigilância popular.</p></a>
      <a class="eixo-card cor-vermelho" href="../eixos/eixo-2/"><small>Eixo II</small><h3>Acesso e integralidade</h3><p>Fluxos de cuidado, direitos e itinerários equânimes.</p></a>
      <a class="eixo-card cor-amarelo" href="../eixos/eixo-3/"><small>Eixo III</small><h3>Comunicação e inovação</h3><p>Letramento digital, informação e justiça climática.</p></a>
    </div>
  </div>
</section>
""",
}

PAGES["eixos/eixo-1/index.html"] = {
    "title": "Eixo 1 · Queremos Viver — PET-Saúde Clima UCPel",
    "current": "eixo-1",
    "body_class": "",
    "main": page_cab(
        "Eixo I",
        "Produção do cuidado no território e vigilância em saúde",
        "Respostas orientadas pela equidade diante de emergências climáticas e ambientais.",
        "cor-azul",
    )
    + """
<section class="secao secao--curta">
  <div class="wrap prosa cor-azul">
    <h2>Objetivos</h2>
    <ul class="lista-check">
      <li>Diagnosticar e intervir em determinantes ambientais e sociais de saúde (insalubridade, umidade e mofo) em domicílios de territórios vulneráveis atingidos por cheias.</li>
      <li>Implementar vigilância popular com mapeamento participativo de riscos sanitários e planejamento de microterritórios saudáveis.</li>
      <li>Qualificar a integração entre vigilância epidemiológica, ambiental e atenção primária, com práticas intersetoriais entre saúde e arquitetura.</li>
    </ul>
    <h2>Grupos de aprendizagem tutorial</h2>
    <p><strong>GAT 1 — Habitat Resiliente e Saúde Ambiental</strong><br>Foco em domicílios, risco sanitário e resiliência habitacional.</p>
    <p><strong>GAT 2 — Sentinelas das Águas: Vigilância e Acessibilidade</strong><br>Vigilância em saúde e acessibilidade em territórios expostos à água.</p>
  </div>
</section>
""",
}

PAGES["eixos/eixo-2/index.html"] = {
    "title": "Eixo 2 · Queremos Viver — PET-Saúde Clima UCPel",
    "current": "eixo-2",
    "body_class": "",
    "main": page_cab(
        "Eixo II",
        "Acesso à atenção especializada e integralidade do cuidado",
        "Garantir direitos e itinerários terapêuticos equânimes após eventos climáticos extremos.",
        "cor-vermelho",
    )
    + """
<section class="secao secao--curta">
  <div class="wrap prosa cor-vermelho">
    <h2>Objetivos</h2>
    <ul class="lista-check">
      <li>Analisar e superar barreiras de acesso à Atenção Especializada para desabrigados e populações impactadas por enchentes.</li>
      <li>Estruturar itinerários terapêuticos equânimes, integrando suporte jurídico e acolhimento psicossocial.</li>
      <li>Qualificar a gestão do cuidado e a comunicação entre os pontos da Rede de Atenção à Saúde (RAS).</li>
    </ul>
    <h2>Grupos de aprendizagem tutorial</h2>
    <p><strong>GAT 3 — Cidadania Plena: Fluxos e Garantia de Direitos</strong><br>Documentação, direitos sociais e matriciamento.</p>
    <p><strong>GAT 4 — Itinerários de Cuidado e Equidade</strong><br>Caminhos assistenciais e práticas interprofissionais.</p>
  </div>
</section>
""",
}

PAGES["eixos/eixo-3/index.html"] = {
    "title": "Eixo 3 · Queremos Viver — PET-Saúde Clima UCPel",
    "current": "eixo-3",
    "body_class": "",
    "main": page_cab(
        "Eixo III",
        "Comunicação e inovação em saúde",
        "Informação acessível, letramento digital e educação permanente frente às emergências climáticas.",
        "cor-amarelo",
    )
    + """
<section class="secao secao--curta">
  <div class="wrap prosa cor-amarelo">
    <h2>Objetivos</h2>
    <ul class="lista-check">
      <li>Desenvolver tecnologias sociais e fluxos informativos digitais contra a desinformação sobre direitos e auxílios.</li>
      <li>Promover letramento digital em saúde e direitos, com ferramentas adaptadas à realidade de Pelotas.</li>
      <li>Apoiar Educação Permanente em Saúde valorizando saberes populares e técnicos.</li>
    </ul>
    <h2>Grupo de aprendizagem tutorial</h2>
    <p><strong>GAT 5 — Conexão e Saber: Letramento Digital e Justiça Climática</strong><br>Comunicação, inovação e acesso democrático à informação.</p>
  </div>
</section>
""",
}

PAGES["equipe/index.html"] = {
    "title": "Equipe · Queremos Viver — PET-Saúde Clima UCPel",
    "current": "equipe",
    "body_class": "",
    "main": page_cab(
        "Equipe",
        "Quem somos",
        "Docentes da UCPel, profissionais da SMS Pelotas e estudantes de seis cursos de graduação.",
        "cor-verde",
    )
    + """
<section class="secao secao--curta">
  <div class="wrap">
    <div class="equipe-bloco">
      <article class="pessoa">
        <strong>Profª Drª Luísa Jardim Corrêa de Oliveira</strong>
        <span>Coordenadora do PET-Saúde Clima UCPel · PPG Saúde no Ciclo Vital</span>
      </article>
    </div>
    <div class="prosa">
      <p>A equipe reúne coordenadores de grupo, tutores, preceptores, orientadores de serviço e estudantes monitores distribuídos em cinco GATs.</p>
      <p>Os números consolidados da seleção e a composição prevista por grupo estão em <a href="em-numeros/">A equipe em números</a>.</p>
      <p>Cursos envolvidos: Medicina, Enfermagem, Psicologia, Odontologia, Direito e Arquitetura, em parceria com a Secretaria Municipal de Saúde de Pelotas.</p>
    </div>
  </div>
</section>
""",
}

# Em números — dados do Edital 044/2026 e resultado final de 08/07/2026
gats = [
    {
        "cor": "cor-azul",
        "eixo": "Eixo I · Cuidado no território",
        "nome": "GAT 1 — Habitat Resiliente e Saúde Ambiental",
        "estudantes": 10,
        "cursos": [("Medicina", 2), ("Enfermagem", 2), ("Psicologia", 2), ("Odontologia", 2), ("Arquitetura", 2)],
    },
    {
        "cor": "cor-azul",
        "eixo": "Eixo I · Cuidado no território",
        "nome": "GAT 2 — Sentinelas das Águas",
        "estudantes": 8,
        "cursos": [("Medicina", 2), ("Enfermagem", 2), ("Odontologia", 2), ("Arquitetura", 2)],
    },
    {
        "cor": "cor-vermelho",
        "eixo": "Eixo II · Acesso e integralidade",
        "nome": "GAT 3 — Cidadania Plena",
        "estudantes": 8,
        "cursos": [("Medicina", 2), ("Psicologia", 2), ("Odontologia", 2), ("Direito", 2)],
    },
    {
        "cor": "cor-vermelho",
        "eixo": "Eixo II · Acesso e integralidade",
        "nome": "GAT 4 — Itinerários de Cuidado e Equidade",
        "estudantes": 8,
        "cursos": [("Medicina", 2), ("Psicologia", 2), ("Enfermagem", 2), ("Direito", 2)],
    },
    {
        "cor": "cor-amarelo",
        "eixo": "Eixo III · Comunicação e inovação",
        "nome": "GAT 5 — Conexão e Saber",
        "estudantes": 8,
        "cursos": [("Medicina", 2), ("Enfermagem", 2), ("Odontologia", 2), ("Direito", 2)],
    },
]

gat_html = []
for g in gats:
    cursos = "".join(f"<li><span>{n}</span><span>{v}</span></li>" for n, v in g["cursos"])
    papéis = barra("Estudantes", g["estudantes"], 10) + barra("Coord./tutor/preceptores*", 4, 10)
    gat_html.append(
        f"""
      <article class="gat-num {g['cor']}">
        <small>{g['eixo']}</small>
        <h3>{g['nome']}</h3>
        <p>{g['estudantes']} vagas de estudantes monitores previstas</p>
        <div class="barras">{papéis}</div>
        <div class="mini-listas">
          <div class="mini-lista">
            <h4>Cursos dos estudantes (previsto)</h4>
            <ul>{cursos}</ul>
          </div>
        </div>
      </article>"""
    )

papeis = [
    ("Coordenadores de grupo", 5, 5),
    ("Tutores", 5, 5),
    ("Estudantes monitores", 38, 40),
    ("Preceptores", 8, 10),
    ("Orientadores de serviço", 0, 3),
]
papeis_barras = "".join(barra(n, a, b) for n, a, b in papeis)
cursos_est = [
    ("Medicina", 10),
    ("Enfermagem", 8),
    ("Psicologia", 6),
    ("Odontologia", 6),
    ("Direito", 6),
    ("Arquitetura", 2),
]
cursos_html = "".join(f"<li><span>{n}</span><span>{v}</span></li>" for n, v in cursos_est)
prec_html = "".join(
    f"<li><span>{n}</span><span>{v}</span></li>"
    for n, v in [
        ("Enfermagem", 4),
        ("Medicina", 1),
        ("Psicologia", 1),
        ("Serviço Social", 1),
        ("Nutrição", 1),
    ]
)

PAGES["equipe/em-numeros/index.html"] = {
    "title": "A equipe em números · Queremos Viver — PET-Saúde Clima UCPel",
    "current": "equipe-numeros",
    "body_class": "",
    "main": page_cab(
        "Equipe",
        "A equipe em números",
        "Composição prevista no edital e resultado da seleção de bolsistas do PET-Saúde Clima UCPel.",
        "cor-verde",
    )
    + f"""
<section class="secao secao--curta">
  <div class="wrap">
    <div class="kpis">
      <div class="kpi cor-azul"><strong>63</strong><span>vagas remuneradas previstas no edital</span></div>
      <div class="kpi cor-verde"><strong>56</strong><span>pessoas selecionadas no resultado final</span></div>
      <div class="kpi cor-amarelo"><strong>38</strong><span>estudantes monitores selecionados (de 40)</span></div>
      <div class="kpi cor-vermelho"><strong>6</strong><span>cursos de graduação representados</span></div>
    </div>

    <h2 class="num-tit">Por grupo de aprendizagem tutorial</h2>
    <div class="gat-nums">
      {''.join(gat_html)}
    </div>
    <p class="nota">* Cada GAT prevê, além dos estudantes, coordenação de grupo, tutoria e preceptoria vinculadas ao território. A distribuição nominal por GAT será atualizada conforme a organização interna da equipe.</p>

    <h2 class="num-tit">Equipe por papel (resultado final · 08/07/2026)</h2>
    <div class="painel" style="margin-bottom:28px">
      <div class="barras">{papeis_barras}</div>
      <p class="nota" style="margin-top:16px;margin-bottom:0">Fonte: resultado final do Edital nº 044/2026 – PRAC/UCPel. Há vagas remanescentes em Arquitetura (estudantes), preceptoria e orientação de serviço.</p>
    </div>

    <div class="painel-duplo">
      <div class="painel">
        <h3>Estudantes monitores por curso</h3>
        <div class="mini-lista"><ul>{cursos_html}</ul></div>
      </div>
      <div class="painel">
        <h3>Preceptores selecionados por formação</h3>
        <div class="mini-lista"><ul>{prec_html}</ul></div>
      </div>
    </div>
  </div>
</section>
""",
}

PAGES["documentos/index.html"] = {
    "title": "Documentos · Queremos Viver — PET-Saúde Clima UCPel",
    "current": "documentos",
    "body_class": "",
    "main": page_cab(
        "O projeto",
        "Documentos",
        "Editais, resultados e materiais públicos do PET-Saúde Clima UCPel.",
    )
    + """
<section class="secao secao--curta">
  <div class="wrap">
    <ul class="doc-lista">
      <li><a href="https://ucpel.edu.br/laravel/public/storage/editais/June2026/yFucvqeewv0EiTqG99pG.pdf" target="_blank" rel="noopener"><strong>Edital nº 044/2026 – PRAC/UCPel</strong><span>Seleção de bolsistas · PDF</span></a></li>
      <li><a href="https://ucpel.edu.br/laravel/public/storage/editais/July2026/OqFsFWDhZvNcvOh8y2kU.pdf" target="_blank" rel="noopener"><strong>Resultado final da seleção</strong><span>Coordenadores, tutores, monitores, preceptores e orientadores · PDF</span></a></li>
      <li><a href="https://ucpel.edu.br/laravel/public/storage/editais/July2026/4dBo5KC5Et41FSC00psy.pdf" target="_blank" rel="noopener"><strong>Edital de vagas remanescentes</strong><span>Complementação de vagas · PDF</span></a></li>
    </ul>
  </div>
</section>
""",
}

PAGES["noticias/index.html"] = {
    "title": "Notícias · Queremos Viver — PET-Saúde Clima UCPel",
    "current": "noticias",
    "body_class": "",
    "main": page_cab(
        "Notícias",
        "Notícias",
        "Cobertura pública sobre o PET-Saúde Clima na UCPel e em Pelotas.",
    )
    + """
<section class="secao secao--curta">
  <div class="wrap">
    <ul class="noticia-lista">
      <li><a href="https://ucpel.edu.br/noticias/conheca-o-projeto-pet-saude-clima-da-ucpel" target="_blank" rel="noopener"><strong>Conheça o projeto PET-Saúde: Clima da UCPel</strong><span>11/08/2026 · Portal UCPel</span></a></li>
      <li><a href="https://ucpel.edu.br/noticias/ucpel-abre-selecao-para-o-pet-saude-clima" target="_blank" rel="noopener"><strong>UCPel abre seleção para o PET-Saúde Clima</strong><span>29/06/2026 · Portal UCPel</span></a></li>
      <li><a href="https://www.radiocom.org.br/2026/08/12/programa-de-saude-climatica-prepara-pelotas-para-possiveis-alagamentos/" target="_blank" rel="noopener"><strong>Programa de saúde climática prepara Pelotas para possíveis alagamentos</strong><span>12/08/2026 · Radiocom 104.5 FM</span></a></li>
      <li><a href="https://ucpel.edu.br/noticias/ucpel-promove-cine-maquetaria-com-foco-nas-mudancas-climaticas" target="_blank" rel="noopener"><strong>Cine Maquetaria com foco nas mudanças climáticas</strong><span>09/10/2026 · Portal UCPel</span></a></li>
    </ul>
  </div>
</section>
""",
}

PAGES["agenda/index.html"] = {
    "title": "Agenda · Queremos Viver — PET-Saúde Clima UCPel",
    "current": "agenda",
    "body_class": "",
    "main": page_cab(
        "Agenda",
        "Agenda",
        "Atividades públicas e marcos do programa. Em atualização com a equipe.",
    )
    + """
<section class="secao secao--curta">
  <div class="wrap prosa">
    <p>O cronograma de campo prioriza o mapeamento de risco em comunidades vulneráveis a alagamentos, em articulação com a Prefeitura de Pelotas e a SMS.</p>
    <p>Novas atividades serão publicadas aqui conforme a programação dos GATs.</p>
    <p class="nota">Sugestão de contribuição: envie datas e materiais à coordenação do PET-Saúde Clima UCPel para inclusão nesta agenda.</p>
  </div>
</section>
""",
}

PAGES["sobre/index.html"] = {
    "title": "Sobre o site · Queremos Viver — PET-Saúde Clima UCPel",
    "current": "sobre",
    "body_class": "",
    "main": page_cab(
        "Sobre",
        "Sobre o site",
        "Portal institucional do projeto Queremos Viver — PET-Saúde Clima UCPel.",
    )
    + """
<section class="secao secao--curta">
  <div class="wrap prosa">
    <p>Este site reúne apresentação, eixos, documentos e a visão consolidada da equipe em números, com base em editais e resultados públicos da UCPel.</p>
    <p>A estrutura foi inspirada no portal do PET-Saúde Clima ClimaPampa (Unipampa), adaptada à identidade e aos dados de Pelotas/UCPel.</p>
    <p>Fontes principais: Edital nº 044/2026 – PRAC/UCPel, resultado final da seleção e notícias do portal da UCPel.</p>
  </div>
</section>
""",
}


def write_page(rel: str, meta: dict) -> None:
    out = ROOT / rel
    out.parent.mkdir(parents=True, exist_ok=True)
    html = shell(
        rel,
        title=meta["title"],
        description=DESC,
        current=meta["current"],
        body_class=meta["body_class"],
        main=meta["main"],
    )
    out.write_text(html, encoding="utf-8")
    print("wrote", rel)


def main() -> None:
    for rel, meta in PAGES.items():
        write_page(rel, meta)


if __name__ == "__main__":
    main()
