/* ============================================================
   NEXUS · campo de profundidade
   Cinco planos fixos atras do conteudo, em perspectiva CSS.
   A profundidade existe parado — o movimento do ponteiro so
   acrescenta paralaxe em quem tem mouse. Num monitor de parede
   a cena continua com camadas; num tablet a paralaxe desliga.

   Uso: <script src="/static/profundidade.js" defer></script>
   Nao depende de nada e nao toca no resto da pagina.
   ============================================================ */
(function () {
  "use strict";

  if (document.getElementById("nexus-profundidade")) return;

  var reduzido = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  var lerParametro = function (nome, alternativa) {
    var valor = getComputedStyle(document.documentElement)
      .getPropertyValue(nome).trim();
    return valor || alternativa;
  };

  // Le as cores do tema: o campo acompanha a paleta sem ser reescrito.
  var acento = lerParametro("--ambar", "#7FD4FF");
  var frio = lerParametro("--frio", "#D99A4E");
  var neon = lerParametro("--neon-rosa", "#FF5FA2");
  var chao = lerParametro("--abismo", "#05070A");

  var palco = document.createElement("div");
  palco.id = "nexus-profundidade";
  palco.setAttribute("aria-hidden", "true");
  palco.style.cssText = [
    "position:fixed", "inset:0", "pointer-events:none", "z-index:-3",
    "overflow:hidden", "perspective:900px", "perspective-origin:50% 42%"
  ].join(";");

  var criar = function (css) {
    var camada = document.createElement("div");
    camada.style.cssText = "position:absolute;inset:-18% -14%;" + css;
    palco.appendChild(camada);
    return camada;
  };

  // 1 · fundo distante: uma banda de luz no horizonte, alta e estreita.
  //     Nao repete os clarões que o tema ja pinta — e um plano de luz
  //     atras das torres, para elas terem contra o que recortar.
  var distante = criar(
    "transform:translateZ(-620px) scale(1.78);" +
    "background:" +
      "linear-gradient(180deg, transparent 34%, " + acento + "2a 52%, " +
        neon + "1e 64%, transparent 78%)"
  );

  // 2 · torres: massas escuras recortadas CONTRA essa banda de luz.
  //     O plano traz a propria luz junto, senao seria preto sobre preto.
  var torres = criar(
    "transform:translateZ(-430px) scale(1.52);" +
    "background:" +
      "repeating-linear-gradient(90deg," +
        "transparent 0 38px," + chao + "f2 38px 62px," +
        "transparent 62px 104px," + chao + "cc 104px 138px," +
        "transparent 138px 196px)," +
      "linear-gradient(180deg, transparent 30%, " + acento + "26 56%, " +
        neon + "16 72%, transparent 90%);" +
    "-webkit-mask-image:linear-gradient(180deg,transparent 24%,#000 54%,#000 88%,transparent 100%);" +
    "mask-image:linear-gradient(180deg,transparent 24%,#000 54%,#000 88%,transparent 100%)"
  );

  // 3 · bruma media: a camada que separa o fundo do conteudo
  var bruma = criar(
    "transform:translateZ(-240px) scale(1.28);" +
    "background:" +
      "linear-gradient(180deg, transparent 40%, " + chao + "66 78%, " + chao + "aa 100%)," +
      "radial-gradient(70vw 40vh at 46% 70%, " + acento + "10 0%, transparent 66%)"
  );

  // 4 · chao em fuga: a grade tecnica correndo sob o conteudo.
  //     Fica bem acima da borda inferior, senao vira um risco na beirada.
  var grade = document.createElement("div");
  grade.style.cssText = [
    "position:absolute", "left:-45%", "right:-45%", "bottom:6%", "height:64%",
    "transform:perspective(64vh) rotateX(72deg)", "transform-origin:bottom center",
    "background:repeating-linear-gradient(90deg," + acento + "3d 0 1px, transparent 1px 74px)," +
      "repeating-linear-gradient(0deg," + acento + "2b 0 1px, transparent 1px 58px)",
    "-webkit-mask-image:linear-gradient(180deg,transparent 0%,#000 46%,#000 96%,transparent 100%)",
    "mask-image:linear-gradient(180deg,transparent 0%,#000 46%,#000 96%,transparent 100%)",
    "opacity:.55"
  ].join(";");
  palco.appendChild(grade);

  // 5 · particulas proximas: poeira/chuva parada, muito fraca
  var perto = criar(
    "transform:translateZ(120px) scale(.9);" +
    "background:" +
      "radial-gradient(1.4px 1.4px at 14% 22%, " + acento + "55, transparent 60%)," +
      "radial-gradient(1.2px 1.2px at 63% 12%, " + neon + "44, transparent 60%)," +
      "radial-gradient(1.6px 1.6px at 38% 68%, " + acento + "3a, transparent 60%)," +
      "radial-gradient(1.2px 1.2px at 82% 74%, " + frio + "44, transparent 60%)," +
      "radial-gradient(1.3px 1.3px at 26% 88%, " + acento + "33, transparent 60%)," +
      "radial-gradient(1.1px 1.1px at 71% 44%, " + neon + "33, transparent 60%);" +
    "background-size:52vw 46vh;opacity:.7"
  );

  // 6 · chuva: riscos finos inclinados, dois planos em profundidades
  //     diferentes. E a assinatura do filme — sem ela a cena fica seca.
  var chuvaLonge = criar(
    "transform:translateZ(-160px) scale(1.18);" +
    "background:repeating-linear-gradient(74deg," +
      "transparent 0 26px," + acento + "14 26px 27px, transparent 27px 44px);" +
    "opacity:.5"
  );
  var chuvaPerto = criar(
    "transform:translateZ(60px) scale(.94);" +
    "background:repeating-linear-gradient(71deg," +
      "transparent 0 52px," + acento + "1f 52px 53.5px, transparent 53.5px 88px);" +
    "opacity:.45"
  );

  document.body.appendChild(palco);

  // ---- paralaxe: so com ponteiro fino, so se o usuario aceita movimento ----
  var fino = window.matchMedia("(pointer: fine)").matches;
  if (reduzido || !fino) return;

  var planos = [
    { no: distante,   forca: 6,  z: -620, escala: 1.78 },
    { no: torres,     forca: 12, z: -430, escala: 1.52 },
    { no: bruma,      forca: 20, z: -240, escala: 1.28 },
    { no: chuvaLonge, forca: 28, z: -160, escala: 1.18 },
    { no: chuvaPerto, forca: 54, z: 60,   escala: 0.94 },
    { no: perto,      forca: 42, z: 120,  escala: 0.9  }
  ];

  var alvoX = 0, alvoY = 0, atualX = 0, atualY = 0, rodando = false;

  var quadro = function () {
    atualX += (alvoX - atualX) * 0.06;
    atualY += (alvoY - atualY) * 0.06;
    for (var i = 0; i < planos.length; i++) {
      var p = planos[i];
      p.no.style.transform =
        "translate3d(" + (atualX * p.forca).toFixed(2) + "px," +
        (atualY * p.forca).toFixed(2) + "px,0) " +
        "translateZ(" + p.z + "px) scale(" + p.escala + ")";
    }
    grade.style.transform =
      "translate3d(" + (atualX * 16).toFixed(2) + "px,0,0) " +
      "perspective(64vh) rotateX(72deg)";

    if (Math.abs(alvoX - atualX) > 0.001 || Math.abs(alvoY - atualY) > 0.001) {
      requestAnimationFrame(quadro);
    } else {
      rodando = false;
    }
  };

  window.addEventListener("pointermove", function (evento) {
    alvoX = (evento.clientX / window.innerWidth - 0.5) * 2;
    alvoY = (evento.clientY / window.innerHeight - 0.5) * 2;
    if (!rodando && !document.hidden) {
      rodando = true;
      requestAnimationFrame(quadro);
    }
  }, { passive: true });
})();
