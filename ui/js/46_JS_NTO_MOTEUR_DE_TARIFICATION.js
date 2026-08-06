// ║  NTO — MOTEUR DE TARIFICATION                                            ║
// ╚══════════════════════════════════════════════════════════════════════════╝

// ── Tranches DGS (bornes supérieures, 13 tranches)
const NTO_DGS_TRANCHES = [9,19,29,39,49,59,69,79,89,99,499,999,3000];
const NTO_DGS_FRAIS    = 2.0;

// ── Grille DGS : [num_dept, nom, t1..t13]
const NTO_DGS = [
[1,"Ain",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[2,"Aisne",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[3,"Allier",23.91885,26.1234,29.03175,31.53645,35.0865,37.84995,42.3522,45.3744,48.5208,51.24285,47.80665,41.22405,38.26395],
[4,"Alpes-de-Haute-Provence",37.4877,41.7933,51.3774,58.57065,63.49725,68.11335,74.8926,81.1233,89.24805,93.34665,88.97895,77.92515,71.9325],
[5,"Hautes-Alpes",37.4877,41.7933,51.3774,58.57065,63.49725,68.11335,74.8926,81.1233,89.24805,93.34665,88.97895,77.92515,71.9325],
[6,"Alpes-Maritimes",34.4448,37.35315,45.59175,47.20635,52.0605,58.1877,66.4056,70.8561,77.1282,81.11295,76.9419,76.31055,73.97145],
[7,"Ardèche",29.2077,34.0515,38.1915,42.13485,48.5415,53.16795,56.8008,64.0044,67.96845,71.6013,67.9167,66.5712,62.7624],
[8,"Ardennes",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[9,"Ariège",34.4448,37.35315,45.59175,47.20635,52.0605,58.1877,66.4056,70.8561,77.1282,81.11295,76.9419,76.31055,73.97145],
[10,"Aube",21.7971,23.0184,25.0677,26.1441,28.6281,29.6631,31.88835,34.60005,35.86275,38.6262,35.80065,27.6345,24.9642],
[11,"Aude",34.4448,37.35315,45.59175,47.20635,52.0605,58.1877,66.4056,70.8561,77.1282,81.11295,76.9419,76.31055,73.97145],
[12,"Aveyron",29.2077,34.0515,38.1915,42.13485,48.5415,53.16795,56.8008,64.0044,67.96845,71.6013,67.9167,66.5712,62.7624],
[13,"Bouches-du-Rhône",29.2077,34.0515,38.1915,42.13485,48.5415,53.16795,56.8008,64.0044,67.96845,71.6013,67.9167,66.5712,62.7624],
[14,"Calvados",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[15,"Cantal",27.87255,31.7745,36.17325,37.80855,42.8283,46.37835,52.89885,56.3868,58.4568,65.0601,61.24095,57.546,54.7308],
[16,"Charente",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[17,"Charente-Maritime",27.87255,31.7745,36.17325,37.80855,42.8283,46.37835,52.89885,56.3868,58.4568,65.0601,61.24095,57.546,54.7308],
[18,"Cher",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[19,"Corrèze",27.87255,31.7745,36.17325,37.80855,42.8283,46.37835,52.89885,56.3868,58.4568,65.0601,61.24095,57.546,54.7308],
[20,"Corse",85.9257,95.1786,102.53745,109.9791,119.6253,123.42375,127.2222,130.95855,134.757,142.3332,142.3332,133.27695,125.3385],
[21,"Côte-d'Or",23.8464,25.5231,27.25155,28.8351,31.70205,33.2856,36.17325,39.3921,41.24475,44.298,41.2137,33.5754,30.9258],
[22,"Côtes-d'Armor",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[23,"Creuse",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[24,"Dordogne",27.87255,31.7745,36.17325,37.80855,42.8283,46.37835,52.89885,56.3868,58.4568,65.0601,61.24095,57.546,54.7308],
[25,"Doubs",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[26,"Drôme",29.2077,34.0515,38.1915,42.13485,48.5415,53.16795,56.8008,64.0044,67.96845,71.6013,67.9167,66.5712,62.7624],
[27,"Eure",21.7971,23.0184,25.0677,26.1441,28.6281,29.6631,31.88835,34.60005,35.86275,38.6262,35.80065,27.6345,24.9642],
[28,"Eure-et-Loir",21.7971,23.0184,25.0677,26.1441,28.6281,29.6631,31.88835,34.60005,35.86275,38.6262,35.80065,27.6345,24.9642],
[29,"Finistère",27.87255,31.7745,36.17325,37.80855,42.8283,46.37835,52.89885,56.3868,58.4568,65.0601,61.24095,57.546,54.7308],
[30,"Gard",29.96325,35.30385,39.5163,45.59175,50.6115,56.29365,63.74565,67.94775,72.27405,75.6792,72.1809,71.23905,68.2065],
[31,"Haute-Garonne",27.87255,31.7745,36.17325,37.80855,42.8283,46.37835,52.89885,56.3868,58.4568,65.0601,61.24095,57.546,54.7308],
[32,"Gers",29.96325,35.30385,39.5163,45.59175,50.6115,56.29365,63.74565,67.94775,72.27405,75.6792,72.1809,71.23905,68.2065],
[33,"Gironde",27.87255,31.7745,36.17325,37.80855,42.8283,46.37835,52.89885,56.3868,58.4568,65.0601,61.24095,57.546,54.7308],
[34,"Hérault",29.96325,35.30385,39.5163,45.59175,50.6115,56.29365,63.74565,67.94775,72.27405,75.6792,72.1809,71.23905,68.2065],
[35,"Ille-et-Vilaine",23.91885,26.1234,29.03175,31.53645,35.0865,37.84995,42.3522,45.3744,48.5208,51.24285,47.80665,41.22405,38.26395],
[36,"Indre",23.8464,25.5231,27.25155,28.8351,31.70205,33.2856,36.17325,39.3921,41.24475,44.298,41.2137,33.5754,30.9258],
[37,"Indre-et-Loire",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[38,"Isère",27.87255,31.7745,36.17325,37.80855,42.8283,46.37835,52.89885,56.3868,58.4568,65.0601,61.24095,57.546,54.7308],
[39,"Jura",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[40,"Landes",29.96325,35.30385,39.5163,45.59175,50.6115,56.29365,63.74565,67.94775,72.27405,75.6792,72.1809,71.23905,68.2065],
[41,"Loir-et-Cher",23.8464,25.5231,27.25155,28.8351,31.70205,33.2856,36.17325,39.3921,41.24475,44.298,41.2137,33.5754,30.9258],
[42,"Loire",27.87255,31.7745,36.17325,37.80855,42.8283,46.37835,52.89885,56.3868,58.4568,65.0601,61.24095,57.546,54.7308],
[43,"Haute-Loire",29.2077,34.0515,38.1915,42.13485,48.5415,53.16795,56.8008,64.0044,67.96845,71.6013,67.9167,66.5712,62.7624],
[44,"Loire-Atlantique",23.91885,26.1234,29.03175,31.53645,35.0865,37.84995,42.3522,45.3744,48.5208,51.24285,47.80665,41.22405,38.26395],
[45,"Loiret",21.7971,23.0184,25.0677,26.1441,28.6281,29.6631,31.88835,34.60005,35.86275,38.6262,35.80065,27.6345,24.9642],
[46,"Lot",29.2077,34.0515,38.1915,42.13485,48.5415,53.16795,56.8008,64.0044,67.96845,71.6013,67.9167,66.5712,62.7624],
[47,"Lot-et-Garonne",29.2077,34.0515,38.1915,42.13485,48.5415,53.16795,56.8008,64.0044,67.96845,71.6013,67.9167,66.5712,62.7624],
[48,"Lozère",29.2077,34.0515,38.1915,42.13485,48.5415,53.16795,56.8008,64.0044,67.96845,71.6013,67.9167,66.5712,62.7624],
[49,"Maine-et-Loire",23.8464,25.5231,27.25155,28.8351,31.70205,33.2856,36.17325,39.3921,41.24475,44.298,41.2137,33.5754,30.9258],
[50,"Manche",23.8464,25.5231,27.25155,28.8351,31.70205,33.2856,36.17325,39.3921,41.24475,44.298,41.2137,33.5754,30.9258],
[51,"Marne",21.7971,23.0184,25.0677,26.1441,28.6281,29.6631,31.88835,34.60005,35.86275,38.6262,35.80065,27.6345,24.9642],
[52,"Haute-Marne",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[53,"Mayenne",23.91885,26.1234,29.03175,31.53645,35.0865,37.84995,42.3522,45.3744,48.5208,51.24285,47.80665,41.22405,38.26395],
[54,"Meurthe-et-Moselle",23.8464,25.5231,27.25155,28.8351,31.70205,33.2856,36.17325,39.3921,41.24475,44.298,41.2137,33.5754,30.9258],
[55,"Meuse",23.91885,26.1234,29.03175,31.53645,35.0865,37.84995,42.3522,45.3744,48.5208,51.24285,47.80665,41.22405,38.26395],
[56,"Morbihan",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[57,"Moselle",23.91885,26.1234,29.03175,31.53645,35.0865,37.84995,42.3522,45.3744,48.5208,51.24285,47.80665,41.22405,38.26395],
[58,"Nièvre",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[59,"Nord",23.8464,25.5231,27.25155,28.8351,31.70205,33.2856,36.17325,39.3921,41.24475,44.298,41.2137,33.5754,30.9258],
[60,"Oise",21.7971,23.0184,25.0677,26.1441,28.6281,29.6631,31.88835,34.60005,35.86275,38.6262,35.80065,27.6345,24.9642],
[61,"Orne",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[62,"Pas-de-Calais",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[63,"Puy-de-Dôme",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[64,"Pyrénées-Atlantiques",34.4448,37.35315,45.59175,47.20635,52.0605,58.1877,66.4056,70.8561,77.1282,81.11295,76.9419,76.31055,73.97145],
[65,"Hautes-Pyrénées",34.4448,37.35315,45.59175,47.20635,52.0605,58.1877,66.4056,70.8561,77.1282,81.11295,76.9419,76.31055,73.97145],
[66,"Pyrénées-Orientales",34.4448,37.35315,45.59175,47.20635,52.0605,58.1877,66.4056,70.8561,77.1282,81.11295,76.9419,76.31055,73.97145],
[67,"Bas-Rhin",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[68,"Haut-Rhin",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[69,"Rhône",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[70,"Haute-Saône",23.91885,26.1234,29.03175,31.53645,35.0865,37.84995,42.3522,45.3744,48.5208,51.24285,47.80665,41.22405,38.26395],
[71,"Saône-et-Loire",23.91885,26.1234,29.03175,31.53645,35.0865,37.84995,42.3522,45.3744,48.5208,51.24285,47.80665,41.22405,38.26395],
[72,"Sarthe",23.8464,25.5231,27.25155,28.8351,31.70205,33.2856,36.17325,39.3921,41.24475,44.298,41.2137,33.5754,30.9258],
[73,"Savoie",29.2077,34.0515,38.1915,42.13485,48.5415,53.16795,56.8008,64.0044,67.96845,71.6013,67.9167,66.5712,62.7624],
[74,"Haute-Savoie",29.96325,35.30385,39.5163,45.59175,50.6115,56.29365,63.74565,67.94775,72.27405,75.6792,72.1809,71.23905,68.2065],
[75,"Paris",29.13525,29.42505,29.5803,29.85975,30.9258,31.2777,31.72275,35.0451,35.28315,38.2536,35.03475,24.37425,20.9898],
[76,"Seine-Maritime",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[77,"Seine-et-Marne",29.13525,29.42505,29.5803,29.85975,30.9258,31.2777,31.72275,35.0451,35.28315,38.2536,35.03475,24.37425,20.9898],
[78,"Yvelines",27.9864,28.51425,28.8972,29.17665,30.5946,31.45365,32.34375,33.58575,34.70355,36.27675,33.12,26.91,20.5137],
[79,"Deux-Sèvres",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[80,"Somme",22.34565,23.8671,25.45065,26.8479,29.3112,30.7188,33.19245,36.31815,37.7568,40.779,37.9431,29.61135,27.58275],
[81,"Tarn",29.96325,35.30385,39.5163,45.59175,50.6115,56.29365,63.74565,67.94775,72.27405,75.6792,72.1809,71.23905,68.2065],
[82,"Tarn-et-Garonne",29.96325,35.30385,39.5163,45.59175,50.6115,56.29365,63.74565,67.94775,72.27405,75.6792,72.1809,71.23905,68.2065],
[83,"Var",29.96325,35.30385,39.5163,45.59175,50.6115,56.29365,63.74565,67.94775,72.27405,75.6792,72.1809,71.23905,68.2065],
[84,"Vaucluse",29.2077,34.0515,38.1915,42.13485,48.5415,53.16795,56.8008,64.0044,67.96845,71.6013,67.9167,66.5712,62.7624],
[85,"Vendée",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[86,"Vienne",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[87,"Haute-Vienne",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[88,"Vosges",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[89,"Yonne",21.7971,23.0184,25.0677,26.1441,28.6281,29.6631,31.88835,34.60005,35.86275,38.6262,35.80065,27.6345,24.9642],
[90,"Territoire de Belfort",23.93955,26.9514,31.41225,33.02685,37.4256,40.20975,44.05995,50.4252,53.2404,57.6288,53.7579,46.48185,44.298],
[91,"Essonne",27.9864,28.51425,28.8972,29.17665,30.5946,31.45365,32.34375,33.58575,34.70355,36.27675,33.12,26.91,20.5137],
[92,"Hauts-de-Seine",27.9864,28.51425,28.8972,29.17665,30.5946,31.45365,32.34375,33.58575,34.70355,36.27675,33.12,26.91,20.5137],
[93,"Seine-Saint-Denis",27.9864,28.51425,28.8972,29.17665,30.5946,31.45365,32.34375,33.58575,34.70355,36.27675,33.12,26.91,20.5137],
[94,"Val-de-Marne",27.9864,28.51425,28.8972,29.17665,30.5946,31.45365,32.34375,33.58575,34.70355,36.27675,33.12,26.91,20.5137],
[95,"Val-d'Oise",29.13525,29.42505,29.5803,29.85975,30.9258,31.2777,31.72275,35.0451,35.28315,38.2536,35.03475,24.37425,20.9898],
[98,"Monaco",41.3379,44.80515,54.69975,56.64555,62.46225,69.8211,79.6743,85.0149,92.58075,97.3107,92.322,91.56645,88.7616]
];

// ── Tranches UTE Cartage IDF (bornes supérieures à partir de l'index 1)
const NTO_UTE_TRANCHES = [0,45,100,200,300,400,500,1000,1500,2000,2500,3000,15000];
// ── Grille UTE : dept -> {rates[13], min}
const NTO_UTE = {
  "60":{rates:[0.32,0.321,0.321,0.321,0.321,0.321,0.1605,0.0749,0.0642,0.0642,0.0535,0.0428,0.0428],min:32.1},
  "75":{rates:[0.16,0.1605,0.1605,0.1605,0.1605,0.1605,0.1177,0.0749,0.0642,0.0642,0.0535,0.0428,0.0428],min:20.33},
  "77":{rates:[0.27,0.2675,0.2675,0.2675,0.2675,0.2675,0.1605,0.0749,0.0642,0.0642,0.0535,0.0428,0.0428],min:24.61},
  "78":{rates:[0.32,0.321,0.321,0.321,0.321,0.321,0.1605,0.0749,0.0642,0.0642,0.0535,0.0428,0.0428],min:32.1},
  "91":{rates:[0.32,0.321,0.321,0.321,0.321,0.321,0.1605,0.0749,0.0642,0.0642,0.0535,0.0428,0.0428],min:32.1},
  "92":{rates:[0.16,0.1605,0.1605,0.1605,0.1605,0.1605,0.107,0.0749,0.0642,0.0642,0.0535,0.0428,0.0428],min:20.33},
  "93":{rates:[0.16,0.1605,0.1605,0.1605,0.1605,0.1605,0.107,0.0749,0.0642,0.0642,0.0535,0.0428,0.0428],min:20.33},
  "94":{rates:[0.16,0.1605,0.1605,0.1605,0.1605,0.1605,0.107,0.0749,0.0642,0.0642,0.0535,0.0428,0.0428],min:20.33},
  "95":{rates:[0.21,0.214,0.214,0.214,0.214,0.214,0.1284,0.0749,0.0642,0.0642,0.0535,0.0428,0.0428],min:21.4}
};

// ── Fuel par mois
const NTO_FUEL = {
"2023-01":0.08,"2023-02":0.08,"2023-03":0.08,"2023-04":0.075,"2023-05":0.07,"2023-06":0.06,
"2023-07":0.06,"2023-08":0.065,"2023-09":0.075,"2023-10":0.085,"2023-11":0.085,"2023-12":0.08,
"2024-01":0.07,"2024-02":0.07,"2024-03":0.075,"2024-04":0.075,"2024-05":0.075,"2024-06":0.07,
"2024-07":0.065,"2024-08":0.065,"2024-09":0.06,"2024-10":0.055,"2024-11":0.055,"2024-12":0.06,
"2025-01":0.06,"2025-02":0.065,"2025-03":0.07,"2025-04":0.065,"2025-05":0.06,"2025-06":0.055,
"2025-07":0.055,"2025-08":0.06,"2025-09":0.06,"2025-10":0.06,"2025-11":0.055,"2025-12":0.065,
"2026-01":0.06,"2026-02":0.165
};

// ── Utilitaires
const IDF_DEPTS = new Set(["60","75","77","78","91","92","93","94","95"]);

function ntoGetFuel(transpType) {
  // Override par transporteur (Options → NTO)
  if (transpType === 'UTE' || transpType === 'Flex') {
    const flexVal = parseFloat(localStorage.getItem('nto_fuel_flex') || '');
    if (!isNaN(flexVal) && flexVal > 0) return flexVal / 100;
  }
  if (transpType === 'DGS') {
    const dgsVal = parseFloat(localStorage.getItem('nto_fuel_dgs') || '');
    if (!isNaN(dgsVal) && dgsVal > 0) return dgsVal / 100;
  }
  // Override global
  const override = parseFloat(localStorage.getItem('nto_fuel_override') || '');
  if (!isNaN(override) && override > 0) return override / 100;
  // Tableau automatique mensuel
  const now = new Date();
  const key = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0');
  return NTO_FUEL[key] !== undefined ? NTO_FUEL[key] : 0.165;
}

function ntoDetectTransp(dept) {
  return IDF_DEPTS.has(String(dept)) ? 'UTE' : 'DGS';
}

// ── Calcul DGS
function ntoCalcDGS(dept, taxable) {
  const deptNum = parseInt(dept);
  const row = NTO_DGS.find(r => r[0] === deptNum);
  if (!row) return null;
  let col = NTO_DGS_TRANCHES.length - 1;
  for (let i = 0; i < NTO_DGS_TRANCHES.length; i++) {
    if (taxable <= NTO_DGS_TRANCHES[i]) { col = i; break; }
  }
  const base      = row[2 + col];
  const avecFrais = base + NTO_DGS_FRAIS;
  const fuel      = ntoGetFuel('DGS');
  const fuelMt    = avecFrais * fuel;
  const total     = avecFrais + fuelMt;
  return { type:'DGS', nom:row[1], dept:deptNum, taxable, tranche:NTO_DGS_TRANCHES[col],
           base, frais:NTO_DGS_FRAIS, avecFrais, fuelRate:fuel, fuelMt, total };
}

// ── Calcul UTE Cartage
function ntoCalcUTE(dept, taxable) {
  const d = NTO_UTE[String(dept)];
  if (!d) return null;
  // tranche : trouver premier i (1..12) où taxable <= TRANCHES[i]
  let rateIdx = 12;
  for (let i = 1; i <= 12; i++) {
    if (taxable <= NTO_UTE_TRANCHES[i]) { rateIdx = i - 1; break; }
  }
  const rate       = d.rates[rateIdx];
  const calculated = taxable * rate;
  const base       = Math.max(d.min, calculated);
  const fuel       = ntoGetFuel('UTE');
  const fuelMt     = base * fuel;
  const total      = base + fuelMt;
  return { type:'UTE', nom:'IDF Dept '+dept, dept, taxable, tranche:NTO_UTE_TRANCHES[rateIdx+1],
           rate, calculated, minimum:d.min, base, isMin:(calculated < d.min),
           fuelRate:fuel, fuelMt, total };
}

// ── Affichage résultat
function ntoAfficher(res) {
  document.getElementById('nto-empty').style.display   = 'none';
  document.getElementById('nto-result').style.display  = 'flex';
  const fmt = v => v.toFixed(2) + ' €';
  const fuelPct = (res.fuelRate * 100).toFixed(1) + '%';
  document.getElementById('nto-res-total').textContent    = fmt(res.total);
  document.getElementById('nto-res-transp-lbl').textContent = res.type === 'DGS' ? 'DGS (13) — National' : 'Cartage / Flex IDF (7)';
  document.getElementById('nto-res-dept-lbl').textContent  = 'Dept ' + res.dept + ' — ' + res.nom + ' · ' + res.taxable + ' kg taxable';

  let rows = '';
  const line = (lbl, val, bold, color) =>
    '<div style="display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:1px solid var(--border)">' +
    '<span style="color:var(--text2)">' + lbl + '</span>' +
    '<span style="font-family:var(--mono);font-weight:' + (bold?'700':'400') + ';color:' + (color||'var(--text)') + '">' + val + '</span></div>';

  if (res.type === 'DGS') {
    rows += line('Poids taxable', res.taxable + ' kg');
    rows += line('Tranche DGS', '≤ ' + res.tranche + ' kg');
    rows += line('Prix de base', fmt(res.base));
    rows += line('Frais de dossier', '+ ' + fmt(res.frais), false, 'var(--orange)');
    rows += line('Sous-total', fmt(res.avecFrais));
    rows += line('Fuel (' + fuelPct + ')', '+ ' + fmt(res.fuelMt), false, 'var(--orange)');
    rows += line('TOTAL TTC', fmt(res.total), true, 'var(--purple)');
  } else {
    rows += line('Poids taxable', res.taxable + ' kg');
    rows += line('Tranche Cartage', '≤ ' + res.tranche + ' kg');
    rows += line('Tarif/kg', res.rate + ' €/kg');
    rows += line('Calculé (' + res.taxable + ' × ' + res.rate + ')', fmt(res.calculated), false, res.isMin ? 'var(--text3)' : 'var(--text)');
    if (res.isMin) rows += line('Minimum garanti', fmt(res.minimum), false, 'var(--orange)');
    rows += line('Base retenue', fmt(res.base));
    rows += line('Fuel (' + fuelPct + ')', '+ ' + fmt(res.fuelMt), false, 'var(--orange)');
    rows += line('TOTAL TTC', fmt(res.total), true, 'var(--purple)');
  }
  document.getElementById('nto-res-detail').innerHTML = rows;

  // Stocker pour copier
  window._ntoLastResult = res;
}

// ── Mode toggle
function ntoSetMode(mode) {
  document.getElementById('nto-panel-j').style.display   = mode === 'j' ? '' : 'none';
  document.getElementById('nto-panel-m').style.display   = mode === 'm' ? '' : 'none';
  document.getElementById('nto-btn-mode-j').className    = 'btn btn-sm ' + (mode === 'j' ? 'btn-purple-out' : 'btn-ghost');
  document.getElementById('nto-btn-mode-m').className    = 'btn btn-sm ' + (mode === 'm' ? 'btn-purple-out' : 'btn-ghost');
  document.getElementById('nto-detected').style.display  = 'none';
  document.getElementById('nto-empty').style.display     = 'flex';
  document.getElementById('nto-result').style.display    = 'none';
}

// ── Charger depuis g_master
function ntoLoadDossier() {
  const j = document.getElementById('nto-inp-j').value.trim();
  if (!j) return toast('Saisir un N° de dossier.');
  const row = g_master.find(r => r.file === j || r.file.includes(j));
  if (!row) return toast('Dossier introuvable dans Dispatch. Vérifiez le N° ou importez d\'abord.');
  const dept    = String(row.dept || '').trim().replace(/^0/, '');
  const poids   = parseFloat(row.poids) || 0;
  const vol     = parseFloat(row.vol)   || 0;
  const taxable = calcTaxable(poids, vol);
  const transp  = ntoDetectTransp(dept);
  // Pré-remplir les inputs éditables
  document.getElementById('nto-det-poids').value = poids.toFixed(2);
  document.getElementById('nto-det-vol').value   = vol.toFixed(3);
  document.getElementById('nto-det-dept').textContent    = dept || '?';
  document.getElementById('nto-det-taxable').textContent = taxable.toFixed(2);
  document.getElementById('nto-det-transp').textContent  = transp === 'UTE' ? 'Cartage / Flex (IDF)' : 'DGS (National)';
  document.getElementById('nto-detected').style.display  = 'block';
  window._ntoAutoData = { dept, taxable, transp };
  toast('Dossier ' + j + ' chargé — ' + poids.toFixed(2) + ' kg brut / ' + vol.toFixed(3) + ' m³ · ' + taxable.toFixed(2) + ' kg taxable');
}

function ntoRecalcFromJ() {
  const dept   = document.getElementById('nto-det-dept').textContent;
  const poids  = parseFloat(document.getElementById('nto-det-poids').value) || 0;
  const vol    = parseFloat(document.getElementById('nto-det-vol').value)   || 0;
  const tax    = calcTaxable(poids, vol);
  const transp = ntoDetectTransp(dept);
  document.getElementById('nto-det-taxable').textContent = tax.toFixed(2);
  window._ntoAutoData = { dept, taxable: tax, transp };
}

// ── Calculer
function ntoCalculer() {
  const modeJ = document.getElementById('nto-panel-j').style.display !== 'none';
  let dept, taxable, forcedTransp;

  if (modeJ) {
    if (!window._ntoAutoData) return toast('Charger d\'abord un dossier.');
    dept         = window._ntoAutoData.dept;
    taxable      = window._ntoAutoData.taxable;
    forcedTransp = window._ntoAutoData.transp;
  } else {
    dept         = document.getElementById('nto-inp-dept').value.trim().replace(/^0/, '');
    taxable      = parseFloat(document.getElementById('nto-inp-taxable').value) || 0;
    forcedTransp = document.getElementById('nto-inp-transp').value;
    if (forcedTransp === 'auto') forcedTransp = ntoDetectTransp(dept);
    if (!dept) return toast('Saisir un département.');
    if (!taxable) return toast('Saisir un poids taxable.');
  }

  const res = forcedTransp === 'UTE' ? ntoCalcUTE(dept, taxable) : ntoCalcDGS(dept, taxable);
  if (!res) return toast('Département ' + dept + ' non trouvé dans la grille. Vérifier le numéro.');
  ntoAfficher(res);
}

// ── Copier
function ntoCopier() {
  const res = window._ntoLastResult;
  if (!res) return;
  const fuelPct = (res.fuelRate * 100).toFixed(1) + '%';
  let txt = '';
  if (res.type === 'DGS') {
    txt = 'NTO DGS — Dept ' + res.dept + ' (' + res.nom + ')\n'
        + 'Poids taxable : ' + res.taxable + ' kg (tranche ≤' + res.tranche + ' kg)\n'
        + 'Prix de base  : ' + res.base.toFixed(2) + ' €\n'
        + 'Frais dossier : +' + res.frais.toFixed(2) + ' €\n'
        + 'Fuel (' + fuelPct + ')  : +' + res.fuelMt.toFixed(2) + ' €\n'
        + '─────────────────────────\n'
        + 'TOTAL         : ' + res.total.toFixed(2) + ' €';
  } else {
    txt = 'NTO Cartage IDF — Dept ' + res.dept + '\n'
        + 'Poids taxable : ' + res.taxable + ' kg (tranche ≤' + res.tranche + ' kg)\n'
        + 'Tarif/kg      : ' + res.rate + ' €/kg\n'
        + 'Calculé       : ' + res.calculated.toFixed(2) + ' €'
        + (res.isMin ? ' → minimum ' + res.minimum.toFixed(2) + ' € appliqué' : '') + '\n'
        + 'Fuel (' + fuelPct + ')  : +' + res.fuelMt.toFixed(2) + ' €\n'
        + '─────────────────────────\n'
        + 'TOTAL         : ' + res.total.toFixed(2) + ' €';
  }
  navigator.clipboard.writeText(txt).then(() => toast('Résultat copié dans le presse-papiers ✓'))
    .catch(() => { const t = document.createElement('textarea'); t.value = txt;
      document.body.appendChild(t); t.select(); document.execCommand('copy');
      document.body.removeChild(t); toast('Résultat copié ✓'); });
}

// ── Init fuel display
function ntoInitFuel() {
  const now  = new Date();
  const key  = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0');
  const tableRate = NTO_FUEL[key] !== undefined ? NTO_FUEL[key] : 0.065;
  const stored = localStorage.getItem('nto_fuel_override');
  const rate   = stored !== null ? parseFloat(stored) / 100 : tableRate;
  document.getElementById('nto-fuel-month').textContent = key;
  const inp = document.getElementById('nto-fuel-rate');
  if (inp) inp.value = (rate * 100).toFixed(1);
}

function ntoSaveFuelOverride() {
  const inp = document.getElementById('nto-fuel-rate');
  if (!inp) return;
  const v = parseFloat(inp.value);
  if (isNaN(v) || v < 0 || v > 30) { toast('Taux fuel invalide (0-30 %).'); return; }
  localStorage.setItem('nto_fuel_override', v.toFixed(1));
  toast('⛽ Fuel mise à jour : ' + v.toFixed(1) + '%');
}


// ╔══════════════════════════════════════════════════════════════════════════╗
