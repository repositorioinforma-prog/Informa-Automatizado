/*
 * Ponte "QWebChannel -> Streamlit".
 *
 * As telas do app desktop Relatória Analítica são páginas HTML/JS que
 * conversavam com o Python pelo QWebChannel do Qt (objeto `QtBridge`).
 * Este arquivo imita essa API dentro de um componente Streamlit, para que o
 * HTML/CSS/JS ORIGINAL rode sem alterações:
 *
 *   QtBridge.metodo(a, b)          -> vira um pedido enviado ao Python
 *   QtBridge.sinal.connect(fn)     -> fn é chamada quando o Python "emite"
 *   page.runJavaScript(codigo)     -> o Python manda código para executar aqui
 *   QMessageBox (avisos)           -> faixa de aviso dentro da página
 *   QFileDialog (abrir arquivo)    -> seletor de arquivo do navegador
 *   QFileDialog (salvar)           -> download pelo navegador
 *
 * Protocolo de componente do Streamlit (v1), sem dependências.
 */
(function () {
  "use strict";

  var NONCE = Math.random().toString(36).slice(2) + Date.now().toString(36);
  var seq = 0;
  var pendentes = [];        // pedidos ainda não confirmados pelo Python
  var ultimoEvento = null;   // último evento do Python já processado
  var pronto = false;
  var timerEnvio = null;
  var sinais = {};           // nome -> [handlers]
  var arquivosPorMetodo = {}; // métodos que abrem seletor de arquivo
  var alturaAplicada = null;
  var alturaAuto = false;     // altura "auto": a área acompanha o conteúdo
  var ultimaAlturaAuto = 0;

  function versaoTela() {
    var m = document.querySelector('meta[name="versao-tela"]');
    return m ? m.getAttribute("content") : "";
  }

  function post(type, extra) {
    var msg = { isStreamlitMessage: true, type: type };
    for (var k in (extra || {})) msg[k] = extra[k];
    window.parent.postMessage(msg, "*");
  }

  function agendarEnvio() {
    if (timerEnvio) return;
    timerEnvio = setTimeout(function () {
      timerEnvio = null;
      if (!pronto || !pendentes.length) return;
      // manda TODOS os pendentes juntos (o Streamlit só guarda o último valor)
      post("streamlit:setComponentValue", {
        value: { nonce: NONCE, pendentes: pendentes.slice(), tela: versaoTela(), visto: ultimoEvento || 0 },
        dataType: "json"
      });
    }, 0);
  }

  function pedir(metodo, args, arquivo) {
    seq += 1;
    var msg = { n: seq, method: metodo, args: args || [] };
    if (arquivo) msg.file = arquivo;
    pendentes.push(msg);
    agendarEnvio();
  }

  // ---------------- seletor de arquivo (no lugar do QFileDialog)
  function escolherArquivo(metodo, args) {
    var cfg = arquivosPorMetodo[metodo] || {};
    var inp = document.createElement("input");
    inp.type = "file";
    if (cfg.accept) inp.accept = cfg.accept;
    inp.style.display = "none";
    inp.addEventListener("change", function () {
      var f = inp.files && inp.files[0];
      if (!f) return;
      avisar("Enviando " + f.name + "…", "info", true);
      var leitor = new FileReader();
      leitor.onload = function () {
        var b64 = String(leitor.result).split(",")[1] || "";
        pedir(metodo, args, { name: f.name, b64: b64 });
      };
      leitor.readAsDataURL(f);
      document.body.removeChild(inp);
    });
    document.body.appendChild(inp);
    inp.click();
  }

  // ---------------- objeto QtBridge
  function canal(nome) {
    var fn = function () {
      var args = Array.prototype.slice.call(arguments);
      if (arquivosPorMetodo[nome]) escolherArquivo(nome, args);
      else pedir(nome, args);
    };
    fn.connect = function (h) { (sinais[nome] = sinais[nome] || []).push(h); };
    fn.disconnect = function () { sinais[nome] = []; };
    return fn;
  }
  var cache = {};
  var ponte = new Proxy({}, {
    get: function (_alvo, nome) {
      if (typeof nome !== "string") return undefined;
      if (nome === "then") return undefined; // não é Promise
      if (!cache[nome]) cache[nome] = canal(nome);
      return cache[nome];
    }
  });

  window.qt = { webChannelTransport: {} };
  window.QWebChannel = function (_transporte, callback) {
    var self = this;
    self.objects = { QtBridge: ponte };
    window.QtBridge = ponte;
    if (typeof callback === "function") callback(self);
  };
  window.QtBridge = ponte;

  // ---------------- área visível (a página rola, não o componente)
  // Devolve {topo, altura}: o trecho deste iframe que está na tela agora.
  function areaVisivel() {
    try {
      var fe = window.frameElement;
      if (fe && window.parent && window.parent.innerHeight) {
        var r = fe.getBoundingClientRect();
        var topo = Math.max(0, -r.top);
        var fim = Math.min(r.height, window.parent.innerHeight - r.top);
        return { topo: topo, altura: Math.max(200, fim - topo) };
      }
    } catch (e) { /* sem acesso à página-mãe */ }
    return { topo: 0, altura: Math.min(window.innerHeight, 900) };
  }
  window.__areaVisivel = areaVisivel;

  var ultimaAlturaJanela = 0;
  function alturaJanela() {
    var h = 860;
    try { h = Math.max(620, (window.parent.innerHeight || 900) - 24); } catch (e) { /* padrão */ }
    if (h !== ultimaAlturaJanela) {
      ultimaAlturaJanela = h;
      post("streamlit:setFrameHeight", { height: h });
    }
  }

  function medirAltura() {
    if (!alturaAuto) return;
    var h = Math.ceil(document.body.getBoundingClientRect().height) + 4;
    if (Math.abs(h - ultimaAlturaAuto) > 1) {
      ultimaAlturaAuto = h;
      post("streamlit:setFrameHeight", { height: h });
    }
  }

  // ---------------- avisos (no lugar do QMessageBox)
  function avisar(texto, nivel, temporario) {
    var cores = {
      info: ["#eef6ff", "#1e5aa8", "#bcd6f5"],
      ok: ["#ecf8ef", "#1f6b35", "#bfe3c8"],
      aviso: ["#fff7e6", "#8a5a00", "#f3d79a"],
      erro: ["#fdecec", "#a12626", "#f2bcbc"]
    }[nivel || "info"] || ["#eef6ff", "#1e5aa8", "#bcd6f5"];
    var caixa = document.getElementById("__avisos_ponte");
    if (!caixa) {
      caixa = document.createElement("div");
      caixa.id = "__avisos_ponte";
      caixa.style.cssText = "position:absolute;top:10px;right:10px;z-index:99999;display:flex;" +
        "flex-direction:column;gap:8px;max-width:440px;font-family:system-ui,Arial,sans-serif";
      document.body.appendChild(caixa);
    }
    caixa.style.top = (areaVisivel().topo + 10) + "px";
    var el = document.createElement("div");
    el.style.cssText = "background:" + cores[0] + ";color:" + cores[1] + ";border:1px solid " + cores[2] +
      ";border-radius:10px;padding:10px 34px 10px 12px;font-size:13px;white-space:pre-wrap;" +
      "box-shadow:0 6px 18px rgba(0,0,0,.10);position:relative";
    el.textContent = texto;
    if (temporario) el.setAttribute("data-temp", "1");
    var x = document.createElement("span");
    x.textContent = "×";
    x.style.cssText = "position:absolute;right:10px;top:6px;cursor:pointer;font-weight:900";
    x.onclick = function () { el.remove(); };
    el.appendChild(x);
    caixa.appendChild(el);
    if (nivel !== "erro") setTimeout(function () { el.remove(); }, 6000);
  }
  window.__avisoPonte = avisar;

  // ---------------- download (no lugar do "Salvar como")
  function baixar(nome, b64, mime) {
    try {
      var bin = atob(b64);
      var bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      var url = URL.createObjectURL(new Blob([bytes], { type: mime || "application/octet-stream" }));
      var a = document.createElement("a");
      a.href = url;
      a.download = nome;
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 4000);
    } catch (e) {
      avisar("Não consegui iniciar o download. Tente de novo; se continuar, verifique se o navegador está bloqueando downloads deste site.", "aviso");
    }
  }

  // ---------------- eventos vindos do Python
  function processar(ev) {
    // resposta do Python chegou: some o "Enviando…"
    var temps = document.querySelectorAll("#__avisos_ponte [data-temp]");
    for (var i = 0; i < temps.length; i++) temps[i].remove();
    try {
      if (ev.tipo === "sinal") {
        (sinais[ev.nome] || []).forEach(function (h) { h.apply(null, ev.args || []); });
      } else if (ev.tipo === "js") {
        (0, eval)(ev.codigo);
      } else if (ev.tipo === "aviso") {
        avisar(ev.texto, ev.nivel);
      } else if (ev.tipo === "download") {
        baixar(ev.nome, ev.b64, ev.mime);
      }
    } catch (e) {
      console.error("Ponte: erro ao processar evento", ev, e);
      avisar("Erro na página: " + e, "erro");
    }
  }

  window.addEventListener("message", function (e) {
    var d = e.data || {};
    if (d.type !== "streamlit:render") return;
    var a = d.args || {};

    if (a.arquivos) arquivosPorMetodo = a.arquivos;
    if (a.altura === "auto") {
      if (!alturaAuto) {
        alturaAuto = true;
        if (window.ResizeObserver) new ResizeObserver(medirAltura).observe(document.body);
        else setInterval(medirAltura, 400);
      }
      medirAltura();
    } else if (a.altura === "janela") {
      // ocupa a altura visível do navegador (como a janela do app desktop)
      if (alturaAplicada !== "janela") {
        alturaAplicada = "janela";
        try { window.parent.addEventListener("resize", alturaJanela); } catch (e) { /* sem acesso */ }
      }
      alturaJanela();
    } else if (a.altura && a.altura !== alturaAplicada) {
      alturaAplicada = a.altura;
      post("streamlit:setFrameHeight", { height: a.altura });
    }

    // confirmações: tira da lista o que o Python já processou
    var ack = (a.acks || {})[NONCE] || 0;
    if (ack) pendentes = pendentes.filter(function (p) { return p.n > ack; });

    var eventos = a.eventos || [];
    if (ultimoEvento === null) {
      // primeira renderização desta página: ignora eventos antigos e pede
      // ao Python o estado atual (dados carregados, resultados, fila)
      ultimoEvento = a.ultimo_evento || 0;
      pronto = true;
      pedir("__init__", []);
    } else {
      eventos.forEach(function (ev) {
        if (ev.n > ultimoEvento) {
          ultimoEvento = ev.n;
          processar(ev);
        }
      });
    }
    agendarEnvio();
  });

  post("streamlit:componentReady", { apiVersion: 1 });
})();
