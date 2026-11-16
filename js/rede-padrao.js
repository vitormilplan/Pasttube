/*
 * Rede de Precedência — Rev02 (16/11/2026), transcrita do PDF.
 * Cada subsistema recebe uma ETAPA (ordem de execução). Etapas menores vêm antes.
 * Edite pela aba "Rede de precedência" do painel; as alterações ficam salvas no navegador
 * e podem ser exportadas/importadas em JSON. Para mudar o padrão de todo mundo, edite este arquivo.
 */
(function (root) {
  var REDE_PADRAO = [
    { subsistema: '421-UTL-PW-502-001', etapa: 1, etapaNome: 'Grupo 2: Hidrantes e água potável', grupo: 'GRUPO 2 - SISTEMAS DE SEGURANÇA',
      descricao: 'LOADING AND STORAGE', predecessoras: '3 - PN-421-MCC-002; 5 - DCS-SIS/FGS',
      sucessoras: '421-UTL-PW-502-002 (EYE WASH SHOWER); 421-UTL-PW-502-003 (SERVICE STATION AND ADMINISTRATIVE)',
      libera: 'Lava-olhos, estação de serviço' },
    { subsistema: '421-UTL-AF-504-003', etapa: 1, etapaNome: 'Grupo 2: Hidrantes e água potável', grupo: 'GRUPO 2 - SISTEMAS DE SEGURANÇA',
      descricao: 'DISTRIBUTION FOR FIRE HYDRANTS', predecessoras: '421-UTL-AF-504-002 (FIRE PUMP SKID)',
      sucessoras: '421-UTL-AF-504-004 (SISTEMA DE GERAÇÃO DE ESPUMA); PLANTA UGNL - REDE DE HIDRANTES',
      libera: 'Espuma, rede de hidrantes' },
    { subsistema: '421-UTL-AF-504-004', etapa: 2, etapaNome: 'Grupo 2: Espuma', grupo: 'GRUPO 2 - SISTEMAS DE SEGURANÇA',
      descricao: 'SISTEMA DE GERAÇÃO DE ESPUMA', predecessoras: '421-UTL-AF-504-003 (DISTRIBUTION FOR FIRE HYDRANTS); 5 - DCS-SIS/FGS',
      sucessoras: '— (FIM DE CADEIA)', libera: 'Fim da cadeia' },
    { subsistema: '421-UTL-AI-503-001', etapa: 3, etapaNome: 'Grupo 3/4: Ar e drenagem', grupo: 'GRUPO 3/4 - UTILIDADES BOP / FURUI',
      descricao: 'AI/AS DISTRIBUTION', predecessoras: '421-PRO-AI-293-002 (INSTRUMENT AIR DISTRIBUTION)',
      sucessoras: '421-UTL-PW-502-003 (SERVICE STATION AND ADMINISTRATIVE)', libera: 'Estação de serviço' },
    { subsistema: '421-UTL-DF-505-001', etapa: 3, etapaNome: 'Grupo 3/4: Ar e drenagem', grupo: 'GRUPO 3/4 - UTILIDADES BOP / FURUI',
      descricao: 'CLOSED DRAINAGE UGNL2', predecessoras: '3 - PN-421-MCC-002; 421-PRO-WW-296-007 (SEWAGE DISCHARGE SYSTEM)',
      sucessoras: '421-PRO-HF-294-001 (FLARE SYSTEM)', libera: 'Flare' },
    { subsistema: '421-PRO-HF-294-001', etapa: 4, etapaNome: 'Grupo 3/4: Flare', grupo: 'GRUPO 3/4 - UTILIDADES BOP / FURUI',
      descricao: 'FLARE SYSTEM (9)', predecessoras: '3 - PN-421-MCC-002; 421-UTL-DF-505-001 (CLOSED DRAINAGE UGNL2); DCS',
      sucessoras: 'Sistemas que aliviam para o flare (nó 9)', libera: 'Sistemas que aliviam no flare' },
    { subsistema: '421-UTL-NG-508-001', etapa: 5, etapaNome: 'Grupo 5/6: Gás natural', grupo: 'GRUPO 5/6 - ENTRADA E COMPRESSÃO DE GÁS NATURAL',
      descricao: 'NATURAL GAS FROM UTG TO UGNL2', predecessoras: '400-PRO-WT-297-001 (DETALHE A); 5 - DCS-SIS/FGS BOP',
      sucessoras: '421-UTL-NG-509-001 (EMED ENEVA)', libera: 'EMED ENEVA' }
  ];
  if (typeof module === 'object' && module.exports) module.exports = REDE_PADRAO;
  else root.REDE_PADRAO = REDE_PADRAO;
})(typeof self !== 'undefined' ? self : this);
