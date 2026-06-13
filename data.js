// ============================================================
// 客戶業績資料檔 — 由 convert_xlsx.py 從業績總表 Excel 自動產生
// 重新產生：python convert_xlsx.py "<Excel檔路徑>"
// ============================================================
//
// 結構說明：
//   每位業代底下依「客戶」拆分，每個客戶再依「產品系列」拆分：
//     channel    : 通路（醫學中心／區域醫院／地區醫院／診所…）
//     system     : 體系名稱   city: 縣市
//     yearTarget : 該客戶該系列年度目標（元）＝ 1~12 月目標金額加總
//     monthly    : 該客戶該系列 1~12 月實際業績（元），尚無資料的月份為 null
//   範圍：處別 Div1/Div2/Div3，不含直營與經銷
//
const CUSTOMER_DATA = {
  year: 2026,
  updatedAt: "2026-05-28",
  salespeople: [
    { name: "張淑華", group: "Div1", customers: [
      { name: "台大醫院", channel: "醫學中心", system: "台大", city: "台北市中正區", products: [
          { series: "Gilenya", yearTarget: 19300680, monthly: [1258744, 1258744, 1888116, 1888116, 1258744, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 4337937, monthly: [326666, 331332, 284666, 321999, 284666, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 4163096, monthly: [349384, 349385, 262038, 349384, 349384, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1716346, monthly: [140668, 139926, 130038, 116500, 145068, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 1611031, monthly: [395520, 131840, 131840, 231741, 395520, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [206664, 162995, 221796, 206664, 221796, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [133144, 133143, 133144, 166430, 166430, null, null, null, null, null, null, null] }
      ] },
      { name: "軍三總", channel: "醫學中心", system: "三總", city: "台北市內湖區", products: [
          { series: "Exelon", yearTarget: 4259326, monthly: [585143, 231168, 281735, 317856, 288961, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1892348, monthly: [223077, 390384, 0, 223076, 501896, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1743067, monthly: [244273, 36279, 125226, 191877, 197788, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 1082847, monthly: [45120, 45119, 0, 45119, 90238, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 618671, monthly: [175872, 43968, 87936, 87936, 87936, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [121429, 157143, 0, 78571, 185714, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [139708, 32146, 37221, 68012, 53166, null, null, null, null, null, null, null] }
      ] },
      { name: "台大醫院癌醫中心", channel: "區域醫院", system: "台大", city: "台北市大安區", products: [
          { series: "Exelon", yearTarget: 424020, monthly: [0, 14000, 28000, 28000, 28000, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 201629, monthly: [15390, 0, 0, 9895, 19418, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 4200, 0, 0, 4200, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1864, 0, 0, 0, 533, null, null, null, null, null, null, null] }
      ] },
      { name: "中國台北", channel: "地區醫院", system: "中國", city: "台北市內湖區", products: [
          { series: "Ritalin LA", yearTarget: 146124, monthly: [17820, 11880, 23760, 11880, 23760, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 89261, monthly: [5434, 0, 8151, 5434, 10868, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 80448, monthly: [8800, 6285, 11313, 12339, 15738, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5623, 1874, 3748, 3748, 5622, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [7428, 0, 2476, 2476, 2476, null, null, null, null, null, null, null] }
      ] },
      { name: "台大醫院北護分院", channel: "地區醫院", system: "台大", city: "台北市萬華區", products: [
          { series: "Ritalin LA", yearTarget: 103122, monthly: [3033, 9099, 12131, 12131, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [6657, 0, 6657, 0, 6657, null, null, null, null, null, null, null] }
      ] },
      { name: "台大金山", channel: "地區醫院", system: "台大", city: "新北市金山區", products: [
          { series: "Ritalin LA", yearTarget: 6673, monthly: [0, 0, 3033, 0, 3033, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "周敏雄", group: "Div1", customers: [
      { name: "高雄長庚", channel: "醫學中心", system: "長庚", city: "高雄市鳥松區", products: [
          { series: "Exelon", yearTarget: 9170702, monthly: [1284074, 0, 715238, 740655, 722865, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 6495191, monthly: [612506, 306252, 1137514, 612506, 393753, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 3552150, monthly: [447415, 0, 195322, 451435, 342622, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 2997271, monthly: [317894, 275700, 408720, 0, 272480, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1925232, monthly: [259080, 0, 141120, 146703, 146582, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [207505, 28353, 132913, 149549, 139007, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [53174, 0, 21280, 22424, 30882, null, null, null, null, null, null, null] }
      ] },
      { name: "高市凱旋", channel: "區域醫院", system: "高市(部標)", city: "高雄市苓雅區", products: [
          { series: "Ritalin LA", yearTarget: 2135550, monthly: [165240, 149524, 106536, 173274, 139969, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 1710933, monthly: [164573, 187787, 246083, 186508, 152004, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 731446, monthly: [92116, 74433, 96202, 85719, 121457, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [28494, 28494, 28494, 21371, 35618, null, null, null, null, null, null, null] }
      ] },
      { name: "阮綜合醫院", channel: "區域醫院", system: "", city: "高雄市苓雅區", products: [
          { series: "Ritalin LA", yearTarget: 1077244, monthly: [137879, 82207, 51819, 165270, 113451, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 674700, monthly: [55467, 55467, 110934, 55467, 55467, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 119557, monthly: [27806, 0, 0, 13760, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [26609, 4628, 13305, 20247, 10991, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7229, 14457, 7229, 7229, 21685, null, null, null, null, null, null, null] }
      ] },
      { name: "部屏東", channel: "區域醫院", system: "部標", city: "屏東縣屏東市", products: [
          { series: "Exelon", yearTarget: 1265291, monthly: [130006, 36112, 72224, 65001, 137227, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 609705, monthly: [38258, 76516, 76516, 38258, 38258, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 170042, monthly: [30096, 0, 15048, 14510, 14510, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14248, 7124, 7124, 14248, 7124, null, null, null, null, null, null, null] }
      ] },
      { name: "高市大同", channel: "區域醫院", system: "長庚", city: "高雄市前金區", products: [
          { series: "Exelon", yearTarget: 851864, monthly: [179446, 42193, 64560, 135726, 101160, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 510799, monthly: [95180, 0, 19380, 72545, 23085, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 132874, monthly: [48304, 6714, 46222, 56450, 12505, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [16467, 671, 9387, 10924, 6695, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [10214, 0, 1915, 3830, 14443, null, null, null, null, null, null, null] }
      ] },
      { name: "高市民生", channel: "地區醫院", system: "高市(部標)", city: "高雄市苓雅區", products: [
          { series: "Exelon", yearTarget: 1126682, monthly: [72226, 72226, 72226, 72226, 72226, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 664598, monthly: [22572, 43532, 66104, 97811, 58042, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [6179, 6179, 6179, 6179, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "屏東基督教", channel: "區域醫院", system: "", city: "屏東縣屏東市", products: [
          { series: "Ritalin LA", yearTarget: 861300, monthly: [65340, 38610, 71280, 71280, 65340, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 298436, monthly: [27907, 5581, 30697, 33488, 16744, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 238800, monthly: [14400, 32400, 0, 17357, 13885, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [33333, 19048, 61905, 21428, 33333, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [10243, 20486, 2049, 21169, 17072, null, null, null, null, null, null, null] }
      ] },
      { name: "屏東榮總", channel: "地區醫院", system: "高榮", city: "屏東縣屏東市", products: [
          { series: "Exelon", yearTarget: 413348, monthly: [103632, 12192, 111494, 22259, 87702, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 92379, monthly: [5280, 0, 11616, 10809, 6109, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 77646, monthly: [12898, 0, 11552, 25026, 3851, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [8281, 1972, 10645, 7124, 5914, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4380, 0, 1251, 6006, 4068, null, null, null, null, null, null, null] }
      ] },
      { name: "屏東榮總龍泉分院", channel: "地區醫院", system: "高榮", city: "屏東縣內埔鄉", products: [
          { series: "Exelon", yearTarget: 443906, monthly: [106176, 0, 26641, 32541, 29080, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 67379, monthly: [32021, 0, 0, 19251, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 62026, monthly: [2640, 7392, 4224, 4074, 6111, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1972, 4732, 2366, 4732, 3154, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 3129, 0, 3754, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部恆春", channel: "地區醫院", system: "部標", city: "屏東縣恆春鎮", products: [
          { series: "Exelon", yearTarget: 520348, monthly: [57784, 57783, 43338, 43338, 57784, null, null, null, null, null, null, null] }
      ] },
      { name: "民眾醫院", channel: "地區醫院", system: "", city: "屏東縣屏東市", products: [
          { series: "Exelon", yearTarget: 526115, monthly: [96693, 20720, 75974, 34533, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3172, 0, 6342, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [228, 0, 228, 988, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "軍高雄802", channel: "區域醫院", system: "軍標", city: "高雄市苓雅區", products: [
          { series: "Exelon Patch", yearTarget: 449841, monthly: [43534, 43534, 29023, 41922, 41924, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 39505, monthly: [9296, 0, 12394, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [9271, 2164, 3400, 9581, 6182, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7857, 0, 0, 3571, 3571, null, null, null, null, null, null, null] }
      ] },
      { name: "高市聯合", channel: "區域醫院", system: "高榮", city: "高雄市鼓山區", products: [
          { series: "Exelon Patch", yearTarget: 258661, monthly: [26394, 21120, 21120, 0, 10183, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 236787, monthly: [25669, 0, 25669, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [11828, 11828, 23657, 0, 11828, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 12514, 0, 12514, null, null, null, null, null, null, null] }
      ] },
      { name: "聖功醫院", channel: "地區醫院", system: "", city: "高雄市苓雅區", products: [
          { series: "Ritalin LA", yearTarget: 251402, monthly: [0, 43474, 0, 57966, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 79623, monthly: [13488, 0, 6744, 0, 6744, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7429, 7429, 0, 14857, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [4286, 2857, 0, 2857, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "南門屏東", channel: "地區醫院", system: "", city: "屏東縣恆春鎮", products: [
          { series: "Exelon", yearTarget: 287297, monthly: [41056, 41057, 13686, 34213, 34214, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3239, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "東港安泰", channel: "地區醫院", system: "安泰", city: "屏東縣東港鎮", products: [
          { series: "Exelon", yearTarget: 64671, monthly: [3000, 6000, 9000, 7200, 4800, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3571, 2857, 7142, 3571, 1429, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 2324, 4648, 4648, 2323, null, null, null, null, null, null, null] }
      ] },
      { name: "輔英醫院", channel: "區域醫院", system: "", city: "屏東縣東港鎮", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [13599, 12305, 13599, 15542, 13599, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [22800, 0, 0, 6080, 10640, null, null, null, null, null, null, null] }
      ] },
      { name: "高市鳳山", channel: "地區醫院", system: "長庚", city: "高雄市鳳山區", products: [
          { series: "Ritalin LA", yearTarget: 45765, monthly: [5493, 0, 10986, 5493, 5493, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 851, 1702, 1702, null, null, null, null, null, null, null] }
      ] },
      { name: "恆春基督教", channel: "地區醫院", system: "", city: "屏東縣恆春鎮", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [2280, 0, 1520, 2280, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "軍高屏815", channel: "地區醫院", system: "軍標", city: "屏東縣屏東市", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [357, 286, 714, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "陳文琪", group: "Div1", customers: [
      { name: "林口長庚", channel: "醫學中心", system: "長庚", city: "桃園市龜山區", products: [
          { series: "Gilenya", yearTarget: 5130259, monthly: [350004, 612505, 175002, 612508, 218753, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 3774694, monthly: [596796, 59476, 225195, 469200, 212489, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1795346, monthly: [189593, 115356, 141352, 268984, 188442, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1513561, monthly: [129680, 126300, 0, 236113, 120439, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [203905, 19483, 126486, 176879, 147413, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 272479, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [18780, 18966, 21546, 19898, 19471, null, null, null, null, null, null, null] }
      ] },
      { name: "北醫雙和", channel: "醫學中心", system: "北醫", city: "新北市中和區", products: [
          { series: "Exelon", yearTarget: 3722264, monthly: [427080, 330614, 305467, 302080, 299586, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 2216035, monthly: [334925, 92606, 231515, 185212, 231515, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 2193696, monthly: [173929, 173011, 272183, 108813, 304676, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 1994719, monthly: [0, 232576, 232577, 139546, 232577, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 766022, monthly: [95752, 47876, 47876, 95753, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [35600, 62552, 89000, 53400, 44752, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14285, 0, 20000, 8571, 8572, null, null, null, null, null, null, null] }
      ] },
      { name: "桃園長庚", channel: "地區醫院", system: "長庚", city: "桃園市龜山區", products: [
          { series: "Exelon", yearTarget: 1692773, monthly: [211471, 0, 127087, 228753, 49816, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1458640, monthly: [141737, 124698, 112965, 106001, 109497, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1232585, monthly: [117720, 107740, 83600, 132248, 107168, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 894265, monthly: [218753, 131251, 131251, 131251, 131251, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [55090, 0, 58209, 39085, 42785, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 136240, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14098, 13699, 14870, 14125, 13965, null, null, null, null, null, null, null] }
      ] },
      { name: "台北長庚", channel: "醫學中心", system: "長庚", city: "台北市松山區", products: [
          { series: "Exelon", yearTarget: 2133080, monthly: [236378, 59476, 168263, 172836, 152504, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1263359, monthly: [120920, 51700, 47300, 142131, 48369, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 988399, monthly: [525005, 0, 0, 87501, 175002, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 635942, monthly: [83456, 57957, 44631, 62256, 46074, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 317892, monthly: [0, 90827, 90827, 0, 90827, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [47229, 3224, 31144, 36065, 34535, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7448, 7448, 7448, 7448, 7448, null, null, null, null, null, null, null] }
      ] },
      { name: "土城長庚", channel: "區域醫院", system: "長庚", city: "新北市土城區", products: [
          { series: "Exelon", yearTarget: 1421513, monthly: [267897, 28468, 40668, 204355, 137253, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1293352, monthly: [106076, 17713, 75684, 189951, 117194, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 441790, monthly: [45900, 0, 20900, 47489, 22245, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 164733, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [59852, 0, 21172, 31009, 29293, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [19205, 0, 7289, 17024, 15747, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "陳均宜", group: "Div1", customers: [
      { name: "彰化基督教", channel: "醫學中心", system: "彰基", city: "彰化縣員林市", products: [
          { series: "Exelon", yearTarget: 13356184, monthly: [1930855, 63307, 1342103, 1152184, 1069884, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 6973812, monthly: [1053326, 0, 595473, 562698, 562012, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 2779962, monthly: [232665, 142972, 282592, 220054, 220054, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 2564639, monthly: [329128, 0, 329128, 141055, 329128, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [310736, 81408, 227236, 190554, 213818, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [39086, 19543, 32572, 39086, 26057, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 0, monthly: [0, 0, 0, 49700, 46410, null, null, null, null, null, null, null] }
      ] },
      { name: "部草屯療", channel: "區域醫院", system: "部標", city: "南投縣草屯鎮", products: [
          { series: "Exelon", yearTarget: 1825807, monthly: [158898, 167978, 143009, 162230, 157800, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1207378, monthly: [82972, 90647, 91927, 90305, 110422, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1006080, monthly: [69973, 89535, 0, 135897, 99208, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [9973, 9973, 9973, 9261, 10685, null, null, null, null, null, null, null] }
      ] },
      { name: "佑民醫院", channel: "地區醫院", system: "", city: "南投縣草屯鎮", products: [
          { series: "Exelon", yearTarget: 887684, monthly: [126613, 51911, 45580, 75967, 58241, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 534006, monthly: [81029, 37804, 13479, 40750, 36416, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 372103, monthly: [24914, 29020, 58683, 0, 29023, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [14371, 4790, 9580, 7186, 11976, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [590, 1179, 2064, 737, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "埔里基督教", channel: "區域醫院", system: "埔基", city: "南投縣埔里鎮", products: [
          { series: "Exelon Patch", yearTarget: 795702, monthly: [54880, 0, 41160, 0, 39694, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 430068, monthly: [18000, 6154, 30619, 0, 21387, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [95759, 5270, 59989, 13892, 76614, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 115498, monthly: [35358, 0, 19643, 7857, 39286, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [12531, 0, 3686, 2949, 3686, null, null, null, null, null, null, null] }
      ] },
      { name: "彰基員林", channel: "地區醫院", system: "彰基", city: "彰化縣員林市", products: [
          { series: "Ritalin LA", yearTarget: 959686, monthly: [115344, 49927, 115344, 114871, 49927, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [13028, 13028, 16285, 1954, 16285, null, null, null, null, null, null, null] }
      ] },
      { name: "彰基雲林", channel: "地區醫院", system: "彰基", city: "雲林縣西螺鎮", products: [
          { series: "Gilenya", yearTarget: 617413, monthly: [47018, 47018, 47018, 94036, 47018, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 97764, monthly: [2302, 9237, 6935, 2302, 8088, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3908, 0, 4560, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "彰基鹿港長青", channel: "地區醫院", system: "彰基", city: "彰化縣鹿港鎮", products: [
          { series: "Ritalin LA", yearTarget: 641624, monthly: [55234, 11512, 55270, 20532, 43692, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [19543, 0, 6514, 10423, 9120, null, null, null, null, null, null, null] }
      ] },
      { name: "埔里榮民", channel: "地區醫院", system: "中榮", city: "南投縣埔里鎮", products: [
          { series: "Exelon Patch", yearTarget: 293901, monthly: [38172, 0, 25447, 24539, 24539, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 115670, monthly: [12923, 0, 12923, 25846, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [6829, 6026, 6026, 6428, 6428, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2442, 0, 2442, 6717, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 11353, monthly: [614, 6144, 1229, 2457, 3072, null, null, null, null, null, null, null] }
      ] },
      { name: "彰基二林", channel: "地區醫院", system: "彰基", city: "彰化縣二林鎮", products: [
          { series: "Ritalin LA", yearTarget: 169055, monthly: [10773, 10773, 10773, 17708, 21546, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [651, 1302, 1302, 1302, 1953, null, null, null, null, null, null, null] }
      ] },
      { name: "部南投", channel: "區域醫院", system: "部標", city: "南投縣南投市", products: [
          { series: "Ritalin LA", yearTarget: 108630, monthly: [12242, 0, 20659, 6886, 13772, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2708, 0, 1710, 4274, 998, null, null, null, null, null, null, null] }
      ] },
      { name: "彰基鹿港", channel: "地區醫院", system: "彰基", city: "彰化縣鹿港鎮", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [34217, 7156, 7580, 19724, 14870, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7817, 1954, 1954, 1954, 3908, null, null, null, null, null, null, null] }
      ] },
      { name: "彰基漢銘", channel: "地區醫院", system: "彰基", city: "彰化縣彰化市", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [10402, 0, 12704, 15006, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 7817, 5211, 5211, 3257, null, null, null, null, null, null, null] }
      ] },
      { name: "彰基南投", channel: "地區醫院", system: "彰基", city: "南投縣南投市", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [326, 326, 0, 522, 652, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "楊智仁", group: "Div1", customers: [
      { name: "成大醫院", channel: "醫學中心", system: "成大", city: "台南市北區", products: [
          { series: "Exelon Patch", yearTarget: 3759234, monthly: [250731, 183051, 343542, 263281, 451462, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 2616589, monthly: [270682, 135341, 0, 315797, 269581, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 1605837, monthly: [122200, 122200, 97760, 195520, 97493, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [307499, 17090, 84770, 178702, 69691, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 615252, monthly: [87892, 0, 43946, 43946, 43946, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [36572, 29258, 65828, 29257, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "奇美永康", channel: "醫學中心", system: "奇美", city: "台南市永康區", products: [
          { series: "Exelon", yearTarget: 3731958, monthly: [306288, 277928, 275092, 323304, 266584, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 3419519, monthly: [237467, 237467, 0, 189974, 379948, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 463303, monthly: [86400, 57600, 86400, 86400, 43200, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 379580, monthly: [32615, 32615, 32615, 32615, 32615, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [17219, 34438, 17219, 34438, 21919, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [11143, 11143, 11143, 0, 13929, null, null, null, null, null, null, null] }
      ] },
      { name: "奇美佳里", channel: "地區醫院", system: "奇美", city: "台南市佳里區", products: [
          { series: "Exelon", yearTarget: 3553774, monthly: [260912, 170160, 243896, 300616, 306288, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 93600, monthly: [14400, 14400, 14400, 14400, 28800, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [8608, 4304, 4304, 8608, 8608, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 24784, monthly: [0, 0, 6523, 6523, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 557, 557, 1114, 1114, null, null, null, null, null, null, null] }
      ] },
      { name: "部台南", channel: "區域醫院", system: "部標", city: "台南市中西區", products: [
          { series: "Exelon", yearTarget: 2558020, monthly: [285295, 114840, 353905, 69336, 176954, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 505613, monthly: [66211, 30096, 40628, 37727, 46042, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [4635, 1854, 7416, 5871, 6179, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1425, 712, 3562, 0, 2850, null, null, null, null, null, null, null] }
      ] },
      { name: "部嘉南療", channel: "區域醫院", system: "部標", city: "台南市仁德區", products: [
          { series: "Ritalin LA", yearTarget: 1324005, monthly: [255686, 0, 116494, 30020, 115130, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 1032616, monthly: [118143, 85640, 92862, 63060, 70281, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 373840, monthly: [59117, 30096, 29021, 56968, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [64112, 0, 14247, 14247, 35618, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [4944, 0, 4944, 0, 4944, null, null, null, null, null, null, null] }
      ] },
      { name: "市立安南", channel: "區域醫院", system: "中國", city: "台南市安南區", products: [
          { series: "Exelon", yearTarget: 1498438, monthly: [217387, 108693, 108694, 163041, 108693, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 434922, monthly: [100572, 0, 37715, 62037, 49177, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 166320, monthly: [17820, 0, 11880, 17820, 17820, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [24762, 0, 9905, 24762, 14857, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1874, 1406, 2812, 2343, 2343, null, null, null, null, null, null, null] }
      ] },
      { name: "台南市立", channel: "區域醫院", system: "秀傳", city: "台南市東區", products: [
          { series: "Exelon", yearTarget: 1208412, monthly: [95558, 70411, 65381, 90528, 90528, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 558369, monthly: [51435, 43236, 45844, 48510, 40402, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 95358, monthly: [22071, 5623, 8851, 12790, -126, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [18666, 7000, 16332, 9332, 13998, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1096, 2190, 0, 3227, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "台南新樓", channel: "區域醫院", system: "新樓", city: "台南市東區", products: [
          { series: "Exelon", yearTarget: 709136, monthly: [105386, 63232, 0, 63232, 126464, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 378184, monthly: [25989, 25989, 25989, 25967, 25989, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 213351, monthly: [16843, 22457, 16843, 33284, 16442, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [14305, 11554, 10344, 19479, 9134, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7772, 2429, 3886, 14572, 6315, null, null, null, null, null, null, null] }
      ] },
      { name: "郭綜合醫院", channel: "地區醫院", system: "郭綜合", city: "台南市中西區", products: [
          { series: "Exelon Patch", yearTarget: 622888, monthly: [200700, 19466, 0, 0, 19226, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 431983, monthly: [147280, 0, 0, 0, 49093, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [14057, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 760, 0, 760, 760, null, null, null, null, null, null, null] }
      ] },
      { name: "台南榮民", channel: "地區醫院", system: "高榮", city: "台南市永康區", products: [
          { series: "Exelon", yearTarget: 677695, monthly: [119253, 119253, 0, 0, 59627, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 213791, monthly: [35636, 0, 15840, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 146308, monthly: [0, 0, 0, 92406, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1251, 1251, null, null, null, null, null, null, null] }
      ] },
      { name: "奇美柳營", channel: "區域醫院", system: "奇美", city: "台南市柳營區", products: [
          { series: "Exelon", yearTarget: 503108, monthly: [28360, 51048, 34032, 56720, 56720, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 255662, monthly: [32615, 0, 0, 32615, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 172800, monthly: [34560, 25920, 43200, 43200, 28800, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [12914, 6457, 12914, 6457, 6457, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [11142, 0, 5571, 16714, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "麻豆新樓", channel: "地區醫院", system: "新樓", city: "台南市麻豆區", products: [
          { series: "Ritalin LA", yearTarget: 290461, monthly: [32486, 0, 97457, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 86463, monthly: [22457, 0, 8982, 8929, 8760, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [16816, 7924, 7924, 7924, 7924, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 41653, monthly: [6323, 0, 6323, 6323, 6323, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4857, 0, 2429, 3400, 4857, null, null, null, null, null, null, null] }
      ] },
      { name: "部新營", channel: "地區醫院", system: "部標", city: "台南市新營區", products: [
          { series: "Exelon Patch", yearTarget: 361324, monthly: [29020, 36276, 14510, 34934, 50304, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2137, 2850, 0, 0, 1425, null, null, null, null, null, null, null] }
      ] },
      { name: "奇美樹林", channel: "醫學中心", system: "奇美", city: "台南市南區", products: [
          { series: "Ritalin LA", yearTarget: 210008, monthly: [0, 0, 0, 46965, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 11143, 0, 5572, null, null, null, null, null, null, null] }
      ] },
      { name: "西門藥局(郭綜合)", channel: "地區醫院", system: "郭綜合", city: "台南市中西區", products: [
          { series: "Exelon Patch", yearTarget: 13139, monthly: [0, 5352, 5272, 0, 2644, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 3600, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 196, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部台南新化", channel: "地區醫院", system: "部標", city: "台南市新化區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [712, 712, 0, 0, 215, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "黃亮熾", group: "Div1", customers: [
      { name: "高醫中和", channel: "醫學中心", system: "高醫", city: "高雄市三民區", products: [
          { series: "Gilenya", yearTarget: 5052723, monthly: [451136, 451136, 0, 406022, 451136, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 3507558, monthly: [331466, 180800, 331466, 253120, 367627, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 2491808, monthly: [127001, 142400, 219600, 165200, 211600, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 900346, monthly: [189546, 47386, 94772, 142158, 94772, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 630029, monthly: [86482, 43241, 43241, 86482, 216206, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [88635, 93845, 134840, 134840, 156055, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [38742, 32286, 38742, 38743, 48428, null, null, null, null, null, null, null] }
      ] },
      { name: "高雄榮總", channel: "醫學中心", system: "高榮", city: "高雄市左營區", products: [
          { series: "Exelon Patch", yearTarget: 1679791, monthly: [117146, 101246, 151760, 150469, 99554, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 1631365, monthly: [184928, 107941, 222310, 298820, 89862, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 1624271, monthly: [135356, 270712, 90238, 225594, 135356, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 243846, monthly: [25669, 51337, 0, 38503, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [33120, 17613, 29706, 33120, 34173, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [18771, 0, 25029, 0, 18771, null, null, null, null, null, null, null] }
      ] },
      { name: "台東基督教", channel: "地區醫院", system: "", city: "台東縣台東市", products: [
          { series: "Exelon", yearTarget: 1559060, monthly: [349200, 0, 75428, 36317, 139680, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 676430, monthly: [77143, 0, 61715, 61715, 138858, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 431190, monthly: [114994, 0, 28748, 56788, 28394, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [32915, 0, 6857, 3429, 11658, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [10629, 0, 0, 16533, 5904, null, null, null, null, null, null, null] }
      ] },
      { name: "高市小港", channel: "區域醫院", system: "高醫", city: "高雄市小港區", products: [
          { series: "Exelon", yearTarget: 1466641, monthly: [150667, 108479, 156695, 174774, 102453, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 627480, monthly: [54000, 32400, 64800, 60914, 43086, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [23668, 17440, 29066, 29066, 29066, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [12914, 12914, 6457, 6457, 19373, null, null, null, null, null, null, null] }
      ] },
      { name: "部旗山", channel: "地區醫院", system: "部標", city: "高雄市旗山區", products: [
          { series: "Exelon", yearTarget: 1438005, monthly: [65004, 151673, 79448, 115561, 151674, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 495720, monthly: [0, 0, 76517, 0, 76517, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 344599, monthly: [21068, 42134, 10534, 30472, 30472, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [12358, 6179, 6179, 12358, 6179, null, null, null, null, null, null, null] }
      ] },
      { name: "馬偕台東", channel: "區域醫院", system: "馬偕", city: "台東縣台東市", products: [
          { series: "Exelon", yearTarget: 1109018, monthly: [160254, 68680, 114467, 45787, 68680, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [43000, 43000, 59091, 43000, 43000, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 209532, monthly: [16543, 22057, 16543, 0, 16543, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 76230, monthly: [10889, 10889, 0, 10486, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [6114, 6114, 6114, 6114, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "台東榮民", channel: "地區醫院", system: "北榮", city: "台東縣台東市", products: [
          { series: "Exelon", yearTarget: 914465, monthly: [0, 0, 0, 86016, 80640, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 602809, monthly: [0, 0, 0, 55697, 61885, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 298383, monthly: [10800, 0, 10800, 19241, 23480, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 9829, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 2160, 1799, 1800, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "義大醫院", channel: "醫學中心", system: "義大", city: "高雄市燕巢區", products: [
          { series: "Exelon", yearTarget: 1206909, monthly: [82809, 90068, 75802, 121027, 111767, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [22301, 35621, 35112, 45620, 30231, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 148629, monthly: [0, 18296, 6861, 22023, 22023, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [15714, 15714, 15714, 15714, 12572, null, null, null, null, null, null, null] }
      ] },
      { name: "高市岡山", channel: "地區醫院", system: "秀傳", city: "高雄市岡山區", products: [
          { series: "Exelon", yearTarget: 1336940, monthly: [100586, 100586, 150879, 100586, 50293, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 149072, monthly: [11554, 23108, 11554, 22697, 21464, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 700, 0, 700, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [122, 0, 0, 0, 122, null, null, null, null, null, null, null] }
      ] },
      { name: "高醫岡山", channel: "地區醫院", system: "高醫", city: "高雄市岡山區", products: [
          { series: "Exelon", yearTarget: 503881, monthly: [57254, 42187, 66292, 81360, 72320, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 329706, monthly: [44000, 33201, 66201, 51857, 26400, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 218618, monthly: [24023, 18017, 24022, 12011, 84080, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [6229, 10380, 6228, 10380, 14533, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3874, 12914, 6457, 14206, 9040, null, null, null, null, null, null, null] }
      ] },
      { name: "義大大昌", channel: "地區醫院", system: "義大", city: "高雄市三民區", products: [
          { series: "Exelon", yearTarget: 762568, monthly: [76781, 37203, 71748, 75773, 37203, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 240093, monthly: [22868, 0, 22868, 22023, 22023, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [12572, 0, 6286, 9429, 9429, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [18438, 0, 1796, 7812, 4609, null, null, null, null, null, null, null] }
      ] },
      { name: "軍左營806", channel: "區域醫院", system: "軍標", city: "高雄市左營區", products: [
          { series: "Exelon", yearTarget: 632943, monthly: [36120, 36120, 72240, 21673, 57793, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 17041, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2857, 0, 0, 2143, 2143, null, null, null, null, null, null, null] }
      ] },
      { name: "義大癌治療", channel: "區域醫院", system: "義大", city: "高雄市燕巢區", products: [
          { series: "Exelon", yearTarget: 379946, monthly: [37203, 65453, 6992, 45253, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 42302, monthly: [10290, 8004, 17152, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [9219, 9219, 0, 9219, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [6286, 0, 3143, 3143, 3143, null, null, null, null, null, null, null] }
      ] },
      { name: "台東聖母", channel: "地區醫院", system: "", city: "台東縣台東市", products: [
          { series: "Exelon", yearTarget: 171303, monthly: [32433, 24324, 16217, 20271, 36489, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 99608, monthly: [0, 14438, 14438, 14437, 14437, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 3909, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [691, 0, 691, 0, 691, null, null, null, null, null, null, null] }
      ] },
      { name: "慈濟關山", channel: "地區醫院", system: "慈濟", city: "台東縣關山鎮", products: [
          { series: "Exelon", yearTarget: 259185, monthly: [25568, 12784, 25568, 25568, 25568, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [9356, 0, 0, 4678, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "高雄秀傳", channel: "地區醫院", system: "秀傳", city: "高雄市路竹區", products: [
          { series: "Exelon", yearTarget: 188409, monthly: [12573, 30176, 5030, 25147, 25147, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 44281, monthly: [12487, 0, 3355, 3231, 6451, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [2800, 0, 4666, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "高市旗津", channel: "地區醫院", system: "高醫", city: "高雄市旗津區", products: [
          { series: "Exelon", yearTarget: 161951, monthly: [36160, 0, 33147, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [8305, 0, 8305, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2583, 0, 2583, 3229, 5812, null, null, null, null, null, null, null] }
      ] },
      { name: "博田國際醫院", channel: "地區醫院", system: "", city: "高雄市左營區", products: [
          { series: "Exelon", yearTarget: 130392, monthly: [28228, 0, 9410, 18542, 18541, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [13143, 0, 6571, 9857, 13143, null, null, null, null, null, null, null] }
      ] },
      { name: "高雄基督教", channel: "地區醫院", system: "", city: "高雄市苓雅區", products: [
          { series: "Exelon", yearTarget: 126315, monthly: [6912, 0, 6912, 3456, 3456, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 2880, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 1886, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "軍岡山814", channel: "地區醫院", system: "軍標", city: "高雄市岡山區", products: [
          { series: "Exelon", yearTarget: 48200, monthly: [11558, 0, 10836, 0, 7224, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 6429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部台東", channel: "地區醫院", system: "部標", city: "台東縣台東市", products: [
          { series: "Exelon Patch", yearTarget: 14780, monthly: [0, 2903, 0, 9889, 5698, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 712, 0, 1425, null, null, null, null, null, null, null] }
      ] },
      { name: "小港安泰", channel: "地區醫院", system: "", city: "高雄市小港區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [1643, 0, 0, 1643, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "曾子駿", group: "Div1", customers: [
      { name: "台中榮總", channel: "醫學中心", system: "中榮", city: "台中市西屯區", products: [
          { series: "Gilenya", yearTarget: 8868006, monthly: [768641, 394133, 1182400, 0, 788266, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 2743675, monthly: [220572, 190464, 239616, 227329, 184321, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1878823, monthly: [125334, 112611, 138058, 120776, 188972, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1815822, monthly: [186084, 139563, 93042, 186084, 186084, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 1313281, monthly: [0, 218880, 175104, 0, 218880, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [151548, 95458, 134677, 134676, 134677, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [54963, 30535, 36644, 67184, 67185, null, null, null, null, null, null, null] }
      ] },
      { name: "光田向上", channel: "區域醫院", system: "光田", city: "台中市沙鹿區", products: [
          { series: "Exelon", yearTarget: 3114908, monthly: [343409, 211418, 301271, 284944, 460684, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 2883219, monthly: [306674, 125177, 292594, 134206, 268412, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 1704738, monthly: [552892, 0, 138223, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1209057, monthly: [152909, 0, 92867, 111279, 136097, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [40808, 45147, 51428, 38070, 51095, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14914, 0, 8400, 10399, 5143, null, null, null, null, null, null, null] }
      ] },
      { name: "童綜合醫院", channel: "區域醫院", system: "童綜合", city: "台中市梧棲區", products: [
          { series: "Exelon", yearTarget: 1147987, monthly: [94668, 74203, 94667, 94667, 74203, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1050230, monthly: [374729, 0, 374729, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 672416, monthly: [116820, 0, 38940, 76954, 38477, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [67371, 0, 22457, 22457, 44914, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [2403, 33330, 0, 17114, 17114, null, null, null, null, null, null, null] }
      ] },
      { name: "為恭醫院", channel: "區域醫院", system: "中國", city: "苗栗縣頭份市", products: [
          { series: "Ritalin LA", yearTarget: 830412, monthly: [171073, 0, 171073, 0, 171073, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 789382, monthly: [119563, 0, 76085, 86955, 38043, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 329449, monthly: [46680, 13337, 0, 39057, 12860, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [11352, 952, 8380, 10323, 8342, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 18743, 9371, 7029, null, null, null, null, null, null, null] }
      ] },
      { name: "大千醫院", channel: "地區醫院", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Exelon Patch", yearTarget: 892734, monthly: [229338, 0, 81882, 117940, 122609, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [110160, 0, 55080, 55080, 55080, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 61170, monthly: [6875, 0, 0, 0, 34371, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3429, 0, 0, 2743, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部苗栗", channel: "區域醫院", system: "部標", city: "苗栗縣苗栗市", products: [
          { series: "Exelon", yearTarget: 518867, monthly: [46947, 36112, 21668, 65001, 36112, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 121892, monthly: [14510, 14510, 17091, 20960, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [8650, 0, 0, 3089, 1854, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1425, 0, 0, 0, 1425, null, null, null, null, null, null, null] }
      ] },
      { name: "通霄光田", channel: "地區醫院", system: "光田", city: "苗栗縣通霄鎮", products: [
          { series: "Exelon Patch", yearTarget: 370056, monthly: [21950, 41198, 1408, 23980, 14504, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 23400, monthly: [1720, 3010, 431, 2581, 1720, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [2314, 771, 2573, 1541, 1799, null, null, null, null, null, null, null] }
      ] },
      { name: "大千竹南", channel: "診所(含門前藥局)", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Ritalin LA", yearTarget: 269422, monthly: [103115, 0, 0, 0, 68743, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [13714, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大千南勢", channel: "地區醫院", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Ritalin LA", yearTarget: 169076, monthly: [0, 49495, 0, 0, 49495, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 13714, 0, 0, 3429, null, null, null, null, null, null, null] }
      ] },
      { name: "大千福安", channel: "診所(含門前藥局)", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Ritalin LA", yearTarget: 161516, monthly: [20623, 0, 54995, 0, 34372, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2057, 0, 2057, 2057, 1371, null, null, null, null, null, null, null] }
      ] },
      { name: "宏仁診所", channel: "診所(含門前藥局)", system: "中國", city: "苗栗縣竹南鎮", products: [
          { series: "Exelon", yearTarget: 131655, monthly: [0, 0, 0, 38043, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 44015, monthly: [0, 0, 13337, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 469, null, null, null, null, null, null, null] }
      ] },
      { name: "李綜合苑裡", channel: "地區醫院", system: "李綜合", city: "苗栗縣苑裡鎮", products: [
          { series: "Exelon Patch", yearTarget: 142014, monthly: [28120, 0, 7017, 13491, 6957, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1768, 1474, 1768, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1520, 0, 0, 1520, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大千大順", channel: "地區醫院", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Ritalin LA", yearTarget: 99659, monthly: [27497, 0, 13748, 0, 27497, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2057, 0, 2057, 0, 4114, null, null, null, null, null, null, null] }
      ] },
      { name: "李綜合大甲", channel: "區域醫院", system: "李綜合", city: "台中市大甲區", products: [
          { series: "Exelon Patch", yearTarget: 99861, monthly: [8421, 0, 14060, 13720, 9738, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [5895, 5895, 5895, 2948, 8843, null, null, null, null, null, null, null] }
      ] },
      { name: "大千崇仁", channel: "診所(含門前藥局)", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Ritalin LA", yearTarget: 36427, monthly: [13748, 13748, 0, 27497, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1371, 686, 0, 686, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大千佑安", channel: "診所(含門前藥局)", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Ritalin LA", yearTarget: 26117, monthly: [3437, 0, 6875, 10312, 6875, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [686, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大眾醫院", channel: "地區醫院", system: "", city: "苗栗縣竹南鎮", products: [
          { series: "Exelon", yearTarget: 27535, monthly: [0, 2899, 0, 0, 6522, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 227, 151, 0, 76, null, null, null, null, null, null, null] }
      ] },
      { name: "弘大醫院", channel: "地區醫院", system: "", city: "苗栗縣苗栗市", products: [
          { series: "Exelon Patch", yearTarget: 18785, monthly: [4696, 4696, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [990, 743, 0, 743, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [691, 346, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大千苗安", channel: "診所(含門前藥局)", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Ritalin LA", yearTarget: 4124, monthly: [4124, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [686, 0, 0, 686, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "美德醫院", channel: "地區醫院", system: "李綜合", city: "台中市大甲區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [2948, 0, 0, 2948, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "簡美淑", group: "Div1", customers: [
      { name: "台北榮總", channel: "醫學中心", system: "北榮本院", city: "台北市北投區", products: [
          { series: "Gilenya", yearTarget: 4803944, monthly: [403693, 605540, 605540, 201847, 605542, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 3191801, monthly: [310730, 174718, 365567, 263425, 260735, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 3187677, monthly: [0, 0, 439680, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 2370387, monthly: [200507, 178228, 222785, 178228, 712919, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1888209, monthly: [227180, 141600, 208803, 149199, 168802, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [211193, 100852, 113502, 113504, 176400, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [34400, 29487, 49143, 54058, 152343, null, null, null, null, null, null, null] }
      ] },
      { name: "慈濟新店", channel: "醫學中心", system: "慈濟", city: "新北市新店區", products: [
          { series: "Exelon", yearTarget: 3259123, monthly: [377119, 62853, 251412, 251412, 188559, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 2502750, monthly: [269250, 54000, 161250, 285000, 162000, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 625560, monthly: [97040, 11091, 92886, 88104, 80095, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [86569, 15103, 41416, 73820, 45753, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [31113, 20743, 31113, 47812, 46671, null, null, null, null, null, null, null] }
      ] },
      { name: "北市關渡", channel: "地區醫院", system: "特殊", city: "台北市北投區", products: [
          { series: "Ritalin LA", yearTarget: 2104338, monthly: [209608, 168144, 212516, 168144, 168144, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 1865466, monthly: [259642, 69401, 171060, 107035, 96283, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 292988, monthly: [32800, 16400, 32800, 10800, 15800, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [50400, 0, 50400, 25200, 25200, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [24572, 24572, 24572, 49144, 14742, null, null, null, null, null, null, null] }
      ] },
      { name: "振興醫院", channel: "區域醫院", system: "", city: "台北市北投區", products: [
          { series: "Ritalin LA", yearTarget: 1749750, monthly: [100800, 100800, 100800, 201600, 100800, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 1286087, monthly: [95336, 84120, 72904, 173848, 86924, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 829620, monthly: [61714, 61714, 61714, 89142, 53485, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14171, 21257, 14171, 14171, 28343, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [16667, 0, 16667, 16667, 16667, null, null, null, null, null, null, null] }
      ] },
      { name: "耕莘新店", channel: "區域醫院", system: "耕莘", city: "新北市新店區", products: [
          { series: "Gilenya", yearTarget: 564850, monthly: [46227, -3358, 0, 47282, -211, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 484597, monthly: [81730, -213, 47081, 45440, 43661, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 362143, monthly: [37791, 37792, 12597, 50388, 25194, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [27022, 41873, 19033, 49862, 19033, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 7143, 7143, 10714, 10713, null, null, null, null, null, null, null] }
      ] },
      { name: "耕莘永和", channel: "地區醫院", system: "耕莘", city: "新北市永和區", products: [
          { series: "Exelon Patch", yearTarget: 391785, monthly: [14680, 5071, 28606, 28817, 14137, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 309111, monthly: [12597, 12597, 31494, 6299, 18896, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [38066, 19033, 19033, 0, 19033, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 7143, 0, 7143, 14284, null, null, null, null, null, null, null] }
      ] },
      { name: "耕莘安康", channel: "區域醫院", system: "耕莘", city: "新北市新店區", products: [
          { series: "Exelon Patch", yearTarget: 224084, monthly: [25559, 13220, 0, 28817, 13051, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [14221, 18405, 12967, 12967, 23843, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 69842, monthly: [12597, 0, 0, 12598, 6299, null, null, null, null, null, null, null] }
      ] },
      { name: "宏濟醫院", channel: "地區醫院", system: "宏慈", city: "新北市新店區", products: [
          { series: "Exelon Patch", yearTarget: 27680, monthly: [0, 0, 7689, 0, 7597, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 21565, monthly: [0, 0, 2897, 0, 2897, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 11570, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "王繼德", group: "Div1", customers: [
      { name: "台大新竹", channel: "醫學中心", system: "台大", city: "新竹市北區", products: [
          { series: "Exelon", yearTarget: 1986424, monthly: [191334, 88667, 140000, 191333, 168001, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 1230504, monthly: [131840, 43946, 43946, 87892, 87892, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 1090908, monthly: [83916, 41958, 83916, 83916, 167834, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 873260, monthly: [64834, 45057, 50372, 77298, 67402, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 550793, monthly: [78855, 48526, 0, 48526, 36394, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [56630, 21566, 44142, 49692, 63566, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [32620, 0, 17975, 29291, 23966, null, null, null, null, null, null, null] }
      ] },
      { name: "部桃療", channel: "區域醫院", system: "部標", city: "桃園市桃園區", products: [
          { series: "Ritalin LA", yearTarget: 3936586, monthly: [882329, 0, 55091, 205008, 667092, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 650069, monthly: [43336, 59947, 64280, 70782, 59947, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [178088, 0, 14247, 73730, 76365, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 64706, monthly: [6019, 9029, 8814, 4352, 4146, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [4325, 7416, 8961, 8652, 8961, null, null, null, null, null, null, null] }
      ] },
      { name: "部桃園", channel: "區域醫院", system: "部標", city: "桃園市桃園區", products: [
          { series: "Exelon", yearTarget: 3027961, monthly: [216676, 72226, 361126, 216676, 288902, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 978120, monthly: [135432, 75240, 58042, 87064, 58042, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 590580, monthly: [55091, 45909, 38258, 76517, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [21629, 21629, 21629, 25954, 21629, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [21371, 14247, 14247, 21371, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "中國新竹", channel: "地區醫院", system: "中國", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 1420211, monthly: [218943, 0, 111936, 209836, 269633, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 1058277, monthly: [157604, 21739, 119563, 92390, 135867, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 389598, monthly: [33343, 51818, 0, 37072, 25720, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [37485, 18742, 9372, 23428, 24365, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [17333, 17333, 17333, 17333, 14857, null, null, null, null, null, null, null] }
      ] },
      { name: "聯新-聯新", channel: "區域醫院", system: "聯新", city: "桃園市平鎮區", products: [
          { series: "Exelon", yearTarget: 1600727, monthly: [119580, 101477, 156331, 150298, 135489, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 1173121, monthly: [90239, 90240, 90239, 90240, 90240, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 135482, monthly: [16492, 4712, 16492, 11369, 7958, null, null, null, null, null, null, null] }
      ] },
      { name: "台大新竹生醫", channel: "區域醫院", system: "台大", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 1115537, monthly: [121314, 0, 181971, 0, 181971, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 878700, monthly: [55999, 51333, 55999, 65332, 55999, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 129452, monthly: [11100, 0, 4944, 12678, 8720, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [43007, 18667, 20937, 23207, 19802, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [13314, 0, 13314, 13314, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "國泰新竹", channel: "地區醫院", system: "國泰", city: "新竹市東區", products: [
          { series: "Exelon Patch", yearTarget: 809147, monthly: [150115, 60045, 180137, 0, 86726, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 494391, monthly: [93640, 26754, 53509, 33443, 40131, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 62799, monthly: [12747, 0, 12747, 0, 9560, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [13451, 0, 7337, 7337, 10394, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [13042, 0, 12567, 5480, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "馬偕新竹", channel: "區域醫院", system: "馬偕", city: "新竹市東區", products: [
          { series: "Exelon", yearTarget: 638156, monthly: [161168, 0, 57690, 83330, 83330, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 444428, monthly: [79405, 0, 0, 79405, 79405, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [57334, -205, 57334, 28667, 28667, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 146896, monthly: [10131, 11036, 11036, 8682, 21487, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [12229, 0, 12229, 22623, 14062, null, null, null, null, null, null, null] }
      ] },
      { name: "軍桃園804", channel: "區域醫院", system: "軍標", city: "桃園市龍潭區", products: [
          { series: "Ritalin LA", yearTarget: 582499, monthly: [167309, 0, 167308, 0, 167308, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [71428, 0, 85715, 0, 71428, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [12361, 0, 9890, 6181, 6181, null, null, null, null, null, null, null] }
      ] },
      { name: "東元綜合醫院", channel: "區域醫院", system: "", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 754857, monthly: [73074, 69486, 95822, 83780, 107872, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [15143, 22714, 7571, 22714, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 40869, monthly: [2725, 8175, 5450, 12113, 4039, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [5128, 0, 1282, 4486, 1922, null, null, null, null, null, null, null] }
      ] },
      { name: "軍新竹813", channel: "地區醫院", system: "軍標", city: "新竹市北區", products: [
          { series: "Ritalin LA", yearTarget: 841216, monthly: [55770, 55770, 34856, 55770, 71262, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [64901, 0, 0, 21633, 6181, null, null, null, null, null, null, null] }
      ] },
      { name: "部桃園(新屋)", channel: "地區醫院", system: "部標", city: "桃園市新屋區", products: [
          { series: "Exelon", yearTarget: 597438, monthly: [50558, 0, 79448, 43335, 36112, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 194119, monthly: [30096, 0, 45144, 43532, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [6179, 0, 12360, 0, 6179, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3562, 4987, 0, 4275, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "天晟醫院(天成生醫)", channel: "區域醫院", system: "天成", city: "桃園市中壢區", products: [
          { series: "Exelon", yearTarget: 264292, monthly: [33170, 19904, 0, 16585, 13270, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [66900, 89200, 0, 22300, 44600, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 127887, monthly: [13605, 13606, 27211, 6719, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新竹榮民", channel: "地區醫院", system: "北榮", city: "新竹縣竹東鎮", products: [
          { series: "Exelon", yearTarget: 213559, monthly: [16128, 16128, 21505, 16128, 16128, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 172800, monthly: [21601, 5400, 16200, 10399, 10402, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [18001, 0, 7200, 10800, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1966, 2457, null, null, null, null, null, null, null] }
      ] },
      { name: "天成醫院(天成生醫)", channel: "地區醫院", system: "天成", city: "桃園市楊梅區", products: [
          { series: "Exelon", yearTarget: 163482, monthly: [9952, 13932, 6634, 19904, 18577, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [24211, 15929, 17840, 20389, 25486, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 9524, monthly: [1361, 1361, 1361, 1344, 4031, null, null, null, null, null, null, null] }
      ] },
      { name: "桃園榮民", channel: "區域醫院", system: "北榮", city: "桃園市桃園區", products: [
          { series: "Exelon", yearTarget: 165549, monthly: [10752, 0, 34944, 0, 10753, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 38077, monthly: [13439, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3932, 0, 2457, 5897, 2457, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1799, 1080, 6480, 3600, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "聯新國際醫院", channel: "區域醫院", system: "聯新", city: "桃園市平鎮區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [81600, 68000, 0, 10200, 68000, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 13618, 10215, 15662, null, null, null, null, null, null, null] }
      ] },
      { name: "聯新-桃新", channel: "地區醫院", system: "聯新", city: "桃園市桃園區", products: [
          { series: "Exelon", yearTarget: 65312, monthly: [0, 5486, 5486, 0, 2743, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 3534, monthly: [2356, 0, 2356, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "天晟醫院", channel: "區域醫院", system: "天成", city: "桃園市中壢區", products: [
          { series: "Ritalin LA", yearTarget: 53321, monthly: [8668, 8669, 0, 4335, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3800, 3800, 0, 2660, 1520, null, null, null, null, null, null, null] }
      ] },
      { name: "天祥醫院(天成生醫)", channel: "地區醫院", system: "天成", city: "桃園市中壢區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [6371, 12743, 0, 0, 6371, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 21768, monthly: [6803, 0, 4082, 13440, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 6129, monthly: [3317, 0, 1327, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "馬偕兒童", channel: "區域醫院", system: "馬偕", city: "新竹市東區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [30424, 0, 14333, 0, 8045, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 3057, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "聯新-平鎮延平藥局", channel: "診所(含門前藥局)", system: "聯新", city: "桃園市平鎮區", products: [
          { series: "Exelon", yearTarget: 15766, monthly: [0, 0, 0, 2194, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "天祥醫院", channel: "地區醫院", system: "天成", city: "桃園市中壢區", products: [
          { series: "Exelon Patch", yearTarget: 9524, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 4087, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [380, 380, 0, 0, 760, null, null, null, null, null, null, null] }
      ] },
      { name: "天成醫院", channel: "地區醫院", system: "天成", city: "桃園市楊梅區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [3040, 3800, 2280, 2280, 3040, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "林欣平", group: "Div1", customers: [
      { name: "中國附醫(急重症)", channel: "醫學中心", system: "中國", city: "台中市北區", products: [
          { series: "Exelon", yearTarget: 4412187, monthly: [423905, 222822, 434774, 326081, 445643, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 2395602, monthly: [256609, 85536, 397385, 0, 171072, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1115086, monthly: [153154, 76960, 116206, 111968, 74143, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [141942, 92458, 171431, 123507, 145142, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 523522, monthly: [0, 87254, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [60915, 32800, 46856, 60915, 56228, null, null, null, null, null, null, null] }
      ] },
      { name: "部台中", channel: "區域醫院", system: "部標", city: "台中市西區", products: [
          { series: "Gilenya", yearTarget: 5233762, monthly: [857256, 180474, 767018, 496305, 270712, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1922700, monthly: [237201, 229550, 198943, 153034, 214246, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 308484, monthly: [67716, 22572, 31010, 42909, 58041, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [42742, 35618, 35619, 64112, 53427, null, null, null, null, null, null, null] }
      ] },
      { name: "慈濟台中", channel: "區域醫院", system: "慈濟", city: "台中市潭子區", products: [
          { series: "Exelon", yearTarget: 2583682, monthly: [184408, 184408, 177056, 184408, 184408, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 562500, monthly: [75000, 0, 90000, 44250, 45000, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 273411, monthly: [39207, 0, 27726, 13349, 26698, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [39306, 37196, 19653, 37196, 19653, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [8298, 4149, 8298, 9335, 8298, null, null, null, null, null, null, null] }
      ] },
      { name: "長安醫院", channel: "地區醫院", system: "", city: "台中市太平區", products: [
          { series: "Exelon", yearTarget: 1459183, monthly: [212832, 35472, 141888, 130064, 94592, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 826826, monthly: [161241, 0, 45348, 102495, 72892, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 180703, monthly: [13158, 4385, 17543, 11843, 13157, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [7500, 3000, 12001, 4500, 8999, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2720, 1360, 3740, 3060, 4080, null, null, null, null, null, null, null] }
      ] },
      { name: "澄清中港", channel: "區域醫院", system: "", city: "台中市西屯區", products: [
          { series: "Gilenya", yearTarget: 857281, monthly: [45120, 135360, 0, 45120, 135360, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 709072, monthly: [109872, 10464, 49704, 88944, 54936, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 250080, monthly: [26256, 21255, 17504, 15953, 22089, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [31143, 12977, 21802, 22319, 18687, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 99684, monthly: [9588, 0, 15979, 6392, 6392, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [6857, 0, 3429, 6858, 5486, null, null, null, null, null, null, null] }
      ] },
      { name: "亞洲附醫", channel: "區域醫院", system: "中國", city: "台中市霧峰區", products: [
          { series: "Exelon Patch", yearTarget: 760628, monthly: [71057, 45531, 64389, 62039, 43879, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 489808, monthly: [48912, 32609, 32608, 43478, 32608, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [21715, 9905, 21715, 0, 19810, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1874, 937, 1874, 2343, 2343, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [-352, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "軍台中803", channel: "區域醫院", system: "軍標", city: "台中市太平區", products: [
          { series: "Ritalin LA", yearTarget: 494195, monthly: [38728, 30983, 30983, 30983, 116185, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 385229, monthly: [37622, 15049, 15048, 29023, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14286, 0, 5714, 15714, 14286, null, null, null, null, null, null, null] }
      ] },
      { name: "大里仁愛", channel: "區域醫院", system: "長庚", city: "台中市大里區", products: [
          { series: "Exelon", yearTarget: 498707, monthly: [65575, 14742, 48292, 48292, 20842, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 127098, monthly: [12180, 4560, 14420, 12832, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 110638, monthly: [14586, 0, 11260, 15372, 2886, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [22345, 19181, 19181, 19181, 22438, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1940, 931, null, null, null, null, null, null, null] }
      ] },
      { name: "老人復健綜合醫院", channel: "地區醫院", system: "中國", city: "台中市北屯區", products: [
          { series: "Exelon", yearTarget: 348667, monthly: [43477, 65217, 76086, 114127, 54347, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 120001, monthly: [38863, 12954, 25907, 69976, 43126, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 59400, monthly: [0, 0, 0, 5940, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [20800, 0, 0, 20800, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 937, 937, 937, null, null, null, null, null, null, null] }
      ] },
      { name: "部豐原", channel: "區域醫院", system: "部標", city: "台中市豐原區", products: [
          { series: "Exelon", yearTarget: 438070, monthly: [72224, 43335, 0, 57781, 43336, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 236529, monthly: [26121, 34828, 17414, 33534, 16767, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [12360, 12360, 18539, 12360, 12360, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2850, 3562, 2850, 8549, 9974, null, null, null, null, null, null, null] }
      ] },
      { name: "中國台中東區", channel: "地區醫院", system: "中國", city: "台中市東區", products: [
          { series: "Exelon", yearTarget: 561739, monthly: [65216, 27173, 38042, 38043, 38043, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 90064, monthly: [6516, 9184, 0, 3631, 6430, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [4952, 4952, 4952, 0, 9904, null, null, null, null, null, null, null] }
      ] },
      { name: "安德生技", channel: "物流", system: "1B", city: "台中市西屯區", products: [
          { series: "Gilenya", yearTarget: 237467, monthly: [94987, 0, 0, 94987, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林新醫院", channel: "區域醫院", system: "林新", city: "台中市南屯區", products: [
          { series: "Exelon Patch", yearTarget: 185294, monthly: [5674, 17143, 17143, 0, 16480, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 92598, monthly: [16900, 0, 8113, 6760, 6760, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [13571, 0, 8143, 5428, 8143, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5000, 0, 1429, 2143, 2857, null, null, null, null, null, null, null] }
      ] },
      { name: "林新醫院烏日", channel: "地區醫院", system: "林新", city: "台中市烏日區", products: [
          { series: "Ritalin LA", yearTarget: 166947, monthly: [0, 0, 27040, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 48801, monthly: [11349, 0, 0, 0, -172, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4286, 0, 0, 2143, 2143, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 5428, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "軍台中803中清院區", channel: "地區醫院", system: "軍標", city: "台中市北區", products: [
          { series: "Exelon Patch", yearTarget: 186595, monthly: [19562, 7525, 15049, 14511, 14511, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1429, 0, 6072, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "益家診所", channel: "診所(含門前藥局)", system: "中國", city: "台中市東區", products: [
          { series: "Ritalin LA", yearTarget: 52866, monthly: [0, 0, 0, 2971, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 26775, monthly: [0, 0, 0, 5434, 2717, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 938, 1406, 938, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1485, 0, 0, 990, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "惠盛醫院", channel: "地區醫院", system: "中國", city: "台中市豐原區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 9904, 0, 9905, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1406, 0, 0, 937, 1406, null, null, null, null, null, null, null] }
      ] },
      { name: "惠和醫院", channel: "地區醫院", system: "中國", city: "南投縣草屯鎮", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 4951, 4952, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [188, 0, 188, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "亞洲附醫豐中分院", channel: "地區醫院", system: "中國", city: "台中市豐原區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [2476, 2476, 1981, 0, 1485, null, null, null, null, null, null, null] }
      ] },
      { name: "澄清平等", channel: "區域醫院", system: "", city: "台中市中區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [1098, 0, 549, 549, 1098, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "陳奕琦", group: "Div1", customers: [
      { name: "嘉義長庚", channel: "區域醫院", system: "長庚", city: "嘉義縣朴子市", products: [
          { series: "Exelon", yearTarget: 1874308, monthly: [341099, 0, 136746, 142337, 140811, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 1788531, monthly: [218752, 87501, 87501, 175002, 175002, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1506356, monthly: [73700, 86640, 141900, 121787, 169628, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1146312, monthly: [212612, 0, 71995, 113384, 99233, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [33750, 35056, 31402, 34844, 37911, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [15614, 0, 7980, 13620, 23115, null, null, null, null, null, null, null] }
      ] },
      { name: "嘉義基督教", channel: "區域醫院", system: "嘉基", city: "嘉義市東區", products: [
          { series: "Gilenya", yearTarget: 1852239, monthly: [142480, 0, 189973, 0, 94987, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 1817355, monthly: [114507, 171760, 114507, 171760, 111072, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1236267, monthly: [0, 131760, 87840, 122000, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 294030, monthly: [12149, 30371, 24297, 24298, 53582, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [45955, 52915, 43190, 52915, 32035, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 9029, 9029, 9029, 9029, null, null, null, null, null, null, null] }
      ] },
      { name: "台大雲林", channel: "區域醫院", system: "台大", city: "雲林縣斗六市", products: [
          { series: "Exelon", yearTarget: 2153636, monthly: [145999, 233333, 74666, 219332, 140000, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 833468, monthly: [60658, 60658, 60658, 121315, 90987, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 419580, monthly: [0, 0, 0, 0, 209791, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [46599, 26628, 39942, 73227, 66570, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 218460, monthly: [30038, 20149, 20149, 19046, 28941, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [27116, 27116, 30899, 30899, 30899, null, null, null, null, null, null, null] }
      ] },
      { name: "慈濟大林", channel: "區域醫院", system: "慈濟", city: "嘉義縣大林鎮", products: [
          { series: "Exelon", yearTarget: 1290484, monthly: [94280, 94280, 47140, 125707, 94280, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1261621, monthly: [89406, 0, 89406, 70501, 80091, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 597750, monthly: [84000, 42000, 40499, 41250, 84000, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [19882, 20726, 19882, 0, 19882, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4149, 4149, 5186, 5186, 10372, null, null, null, null, null, null, null] }
      ] },
      { name: "嘉義榮民", channel: "區域醫院", system: "中榮", city: "嘉義市西區", products: [
          { series: "Exelon", yearTarget: 2046251, monthly: [92160, 122880, 122880, 92160, 122880, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 738391, monthly: [116302, 32307, 77536, 64614, 71075, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 190845, monthly: [0, 25447, 38172, 0, -2488, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [14060, 28120, 14060, 14060, 14060, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [6108, 6108, 6108, 6108, 18324, null, null, null, null, null, null, null] }
      ] },
      { name: "聖馬爾定", channel: "區域醫院", system: "", city: "嘉義市西區", products: [
          { series: "Exelon", yearTarget: 1306538, monthly: [114453, 0, 171680, 28613, 85840, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 921457, monthly: [100103, 74811, 87545, 36326, 48023, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [66666, 0, 33333, 66666, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4300, 0, 2457, 6143, 2457, null, null, null, null, null, null, null] }
      ] },
      { name: "嘉義陽明", channel: "地區醫院", system: "", city: "嘉義市西區", products: [
          { series: "Exelon", yearTarget: 988934, monthly: [170000, 0, 113334, 113334, 113334, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 26524, 26524, 26524, 26524, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1140, 0, 0, 380, 1140, null, null, null, null, null, null, null] }
      ] },
      { name: "慈濟斗六", channel: "地區醫院", system: "慈濟", city: "雲林縣斗六市", products: [
          { series: "Exelon Patch", yearTarget: 335486, monthly: [40594, 7176, 18267, 35210, 13838, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 322624, monthly: [31426, 18856, 18856, 26397, 25140, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 45750, monthly: [7500, 0, 7502, 3751, 7500, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [2339, 7017, 0, 0, 7017, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [519, 519, 519, 519, 1556, null, null, null, null, null, null, null] }
      ] },
      { name: "部嘉義", channel: "地區醫院", system: "部標", city: "嘉義市東區", products: [
          { series: "Exelon", yearTarget: 307616, monthly: [43335, 0, 10834, 11555, 23110, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 134952, monthly: [23217, 0, 0, 6986, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [8650, 1854, 10504, 2472, 4325, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1069, 0, 712, 1424, null, null, null, null, null, null, null] }
      ] },
      { name: "慈濟嘉義", channel: "診所(含門前藥局)", system: "慈濟", city: "嘉義市西區", products: [
          { series: "Exelon Patch", yearTarget: 203146, monthly: [23714, 14010, 19408, 13497, 18690, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 51621, monthly: [0, 6286, 6286, 6286, 6286, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1754, 1754, 1754, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "灣橋榮民", channel: "地區醫院", system: "中榮", city: "嘉義縣竹崎鄉", products: [
          { series: "Exelon", yearTarget: 209986, monthly: [18433, 30721, 18433, 0, 24576, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [16068, 0, 4018, 0, 4018, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 14216, monthly: [0, 3878, 5169, 0, 1938, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [611, 916, 611, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 0, monthly: [0, 0, -424, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "雲林長庚", channel: "地區醫院", system: "長庚", city: "雲林縣麥寮鄉", products: [
          { series: "Exelon", yearTarget: 145019, monthly: [6100, 11692, 4067, 4576, 20335, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [7238, 6152, 5790, 7781, 6515, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 29900, monthly: [3661, 0, 4272, 3661, 4272, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 19317, monthly: [3420, 3300, 0, 0, 9614, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [373, 452, 532, 0, 319, null, null, null, null, null, null, null] }
      ] },
      { name: "慶昇醫院", channel: "地區醫院", system: "", city: "嘉義市西區", products: [
          { series: "Exelon", yearTarget: 157713, monthly: [7317, 7317, 7318, 13171, 13171, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [8094, 6477, 1942, 7769, 4857, null, null, null, null, null, null, null] }
      ] },
      { name: "成大斗六", channel: "地區醫院", system: "成大", city: "雲林縣斗六市", products: [
          { series: "Exelon", yearTarget: 67246, monthly: [9776, 0, 9776, 0, 10737, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 51888, monthly: [2256, 3384, 4512, 3261, 6275, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [12085, 0, 5371, 5371, 5295, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3291, 0, 1463, 3218, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部朴子", channel: "地區醫院", system: "部標", city: "嘉義縣朴子市", products: [
          { series: "Exelon", yearTarget: 123049, monthly: [14446, 0, 10834, 14446, 7223, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 10158, monthly: [0, 0, 7256, 5429, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [620, 1546, 1545, 1236, 1854, null, null, null, null, null, null, null] }
      ] },
      { name: "台大虎尾", channel: "區域醫院", system: "台大", city: "雲林縣虎尾鎮", products: [
          { series: "Gilenya", yearTarget: 83916, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "信安醫院", channel: "地區醫院", system: "", city: "雲林縣斗六市", products: [
          { series: "Exelon", yearTarget: 12063, monthly: [0, 2074, 3456, 4838, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 6972, monthly: [0, 0, 0, 2766, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "李盈瑩", group: "Div1", customers: [
      { name: "北醫萬芳", channel: "醫學中心", system: "北醫", city: "台北市文山區", products: [
          { series: "Exelon", yearTarget: 4094209, monthly: [468601, 63840, 408520, 344679, 282626, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1839222, monthly: [134258, 76286, 240445, 144071, 144071, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1306319, monthly: [111100, 111126, 111126, 0, 111126, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 617413, monthly: [139546, 46515, 93031, 139546, 46515, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [47040, 44752, 17800, 26952, 35600, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [17143, 17143, 17143, 17142, 17143, null, null, null, null, null, null, null] }
      ] },
      { name: "輔大附醫", channel: "區域醫院", system: "", city: "新北市泰山區", products: [
          { series: "Ritalin LA", yearTarget: 3720536, monthly: [382794, 191397, 382794, 382792, 382792, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 3162538, monthly: [206240, 178704, 285280, 240576, 120288, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 810707, monthly: [61560, 49575, 61560, 97932, 23690, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [71097, 34721, 69440, 69440, 69440, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [40971, 40971, 40971, 61457, 40972, null, null, null, null, null, null, null] }
      ] },
      { name: "新光醫院", channel: "醫學中心", system: "新光", city: "台北市士林區", products: [
          { series: "Exelon", yearTarget: 2295483, monthly: [157998, 184334, 236997, 131665, 157998, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1045166, monthly: [121272, 40424, 80848, 80848, 80104, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [14742, 58972, 117944, 0, 117942, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [44800, 47228, 44800, 47228, 47228, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [24114, 0, 24114, 24114, 24114, null, null, null, null, null, null, null] }
      ] },
      { name: "國泰汐止", channel: "區域醫院", system: "國泰", city: "新北市汐止區", products: [
          { series: "Ritalin LA", yearTarget: 1941438, monthly: [385261, 0, 288946, 337104, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 539676, monthly: [95600, 0, 44613, 44613, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 330264, monthly: [30023, 15012, 60046, 28908, 28908, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [12229, 12229, 12229, 36686, 12229, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [37186, 4739, 16440, 16440, 2369, null, null, null, null, null, null, null] }
      ] },
      { name: "國泰總院", channel: "醫學中心", system: "國泰", city: "台北市大安區", products: [
          { series: "Exelon Patch", yearTarget: 840672, monthly: [135105, 0, 90070, 86724, 86724, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 802800, monthly: [163202, 0, 96316, 48158, 48158, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 173351, monthly: [19120, 0, 12747, 19120, 19120, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [27400, 0, 11743, 18851, 11743, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [12229, 0, 12229, 12229, 18342, null, null, null, null, null, null, null] }
      ] },
      { name: "中心綜合醫院", channel: "地區醫院", system: "", city: "台北市大安區", products: [
          { series: "Exelon Patch", yearTarget: 713025, monthly: [69299, 45357, 90714, 0, 29304, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 418552, monthly: [41816, 17248, 41815, 26134, 26133, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1262, 2524, 5048, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部台北", channel: "區域醫院", system: "部標", city: "新北市新莊區", products: [
          { series: "Ritalin LA", yearTarget: 849150, monthly: [98706, 76517, 76516, 76515, 122425, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [31345, 0, 19235, 14248, 26359, null, null, null, null, null, null, null] }
      ] },
      { name: "景美醫院", channel: "地區醫院", system: "佳醫集康", city: "台北市文山區", products: [
          { series: "Ritalin LA", yearTarget: 347976, monthly: [31967, 33480, 34512, 31967, 25368, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 32025, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "信光藥局(新光醫院)", channel: "健保藥局", system: "1B", city: "台北市士林區", products: [
          { series: "Ritalin LA", yearTarget: 367970, monthly: [49388, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市陽明", channel: "區域醫院", system: "北市", city: "台北市士林區", products: [
          { series: "Exelon", yearTarget: 129667, monthly: [10828, 7219, 7219, 7219, 18047, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [24686, 12343, 9257, 15429, 12343, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 49307, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 7143, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "佳赫藥局(新光醫院)", channel: "健保藥局", system: "1B", city: "台北市士林區", products: [
          { series: "Ritalin LA", yearTarget: 254369, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "綠杏萬芳藥局(北醫萬芳)", channel: "健保藥局", system: "1B", city: "台北市文山區", products: [
          { series: "Ritalin LA", yearTarget: 46922, monthly: [8232, 0, 4116, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "佳醫-景美醫院", channel: "地區醫院", system: "佳醫集康", city: "台北市文山區", products: [
          { series: "Exelon", yearTarget: 32025, monthly: [6934, 1385, 3467, 0, 3467, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3157, 0, 2211, 1263, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "國泰內湖", channel: "診所(含門前藥局)", system: "國泰", city: "台北市內湖區", products: [
          { series: "Ritalin LA", yearTarget: 19401, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 783, 391, 1565, 1174, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [611, 0, 611, 917, 611, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "柯泓瑋", group: "Div1", customers: [
      { name: "台中中山附醫", channel: "醫學中心", system: "台中中山", city: "台中市南區", products: [
          { series: "Exelon", yearTarget: 2785838, monthly: [378314, 91317, 221766, 313086, 260906, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1615572, monthly: [227115, 41700, 125118, 123711, 155125, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 1198732, monthly: [138315, 0, 92211, 92211, 92211, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 1027491, monthly: [177600, 88800, 88800, 0, 88800, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 623662, monthly: [101828, 152742, 0, 50914, 101828, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [121263, 20210, 62507, 66262, 80842, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [24948, 10572, 12158, 10572, 2643, null, null, null, null, null, null, null] }
      ] },
      { name: "彰濱秀傳", channel: "區域醫院", system: "秀傳", city: "彰化縣鹿港鎮", products: [
          { series: "Exelon", yearTarget: 2165438, monthly: [377199, 0, 100587, 100586, 150880, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 947521, monthly: [360960, 0, 0, 90240, 90240, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 622069, monthly: [154903, 0, 61961, 123922, 61961, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 496634, monthly: [96873, 0, 43235, 55498, 39657, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [9332, 18666, 4666, 4666, 18664, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7302, 0, 2434, 3651, 3651, null, null, null, null, null, null, null] }
      ] },
      { name: "彰化秀傳", channel: "區域醫院", system: "秀傳", city: "彰化縣彰化市", products: [
          { series: "Exelon", yearTarget: 2136014, monthly: [226318, 90529, 211230, 170997, 145849, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 964635, monthly: [102127, 56843, 79390, 76513, 73640, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 676801, monthly: [45120, 45120, 45120, 135360, 45120, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 128415, monthly: [12910, 12908, 12910, 12910, 12910, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [20533, 8400, 14000, 19600, 11200, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3043, 4869, 0, 4869, 4869, null, null, null, null, null, null, null] }
      ] },
      { name: "中國北港", channel: "區域醫院", system: "中國", city: "雲林縣北港鎮", products: [
          { series: "Exelon", yearTarget: 1497880, monthly: [391296, 54347, 163040, 130432, 27173, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 517514, monthly: [53349, 40011, 26674, 64300, 51440, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 87912, monthly: [23760, 0, 0, 0, 23760, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [4952, 0, 4952, 0, 4952, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2343, 0, 1406, 2812, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "若瑟醫院", channel: "地區醫院", system: "", city: "雲林縣虎尾鎮", products: [
          { series: "Exelon", yearTarget: 1249818, monthly: [139667, 66817, 224587, 0, 54747, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 442242, monthly: [44297, 73828, 110743, 0, 47988, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 265547, monthly: [51919, 22251, 44503, 29668, 14834, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [39503, 0, 32614, 15833, 25334, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [8675, 10843, 7228, 7228, 13012, null, null, null, null, null, null, null] }
      ] },
      { name: "宏仁醫院", channel: "地區醫院", system: "宏仁", city: "彰化縣員林市", products: [
          { series: "Exelon", yearTarget: 2154960, monthly: [278400, 139200, 139200, 69600, 208800, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [7963, 1594, 7009, 3186, 6372, null, null, null, null, null, null, null] }
      ] },
      { name: "部彰化", channel: "區域醫院", system: "部標", city: "彰化縣埔心鄉", products: [
          { series: "Ritalin LA", yearTarget: 1177335, monthly: [153033, 153033, 76517, 229550, 153033, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 189605, monthly: [34610, 15048, 7524, 23217, 29020, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 49865, 0, 14247, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "竹山秀傳", channel: "地區醫院", system: "秀傳", city: "南投縣竹山鎮", products: [
          { series: "Exelon", yearTarget: 789779, monthly: [120704, 20117, 45265, 55325, 30176, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 451771, monthly: [79393, 0, 16773, 54600, 37455, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [23333, 0, 9333, 9333, 9333, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 50979, monthly: [12910, 0, 12908, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5477, 0, 913, 3652, 913, null, null, null, null, null, null, null] }
      ] },
      { name: "東華醫院", channel: "地區醫院", system: "宏仁", city: "南投縣竹山鎮", products: [
          { series: "Exelon", yearTarget: 968157, monthly: [200449, 0, 0, 27840, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 48202, monthly: [22594, 0, 0, 0, 30126, null, null, null, null, null, null, null] }
      ] },
      { name: "員榮員生", channel: "地區醫院", system: "員榮", city: "彰化縣員林市", products: [
          { series: "Exelon", yearTarget: 249334, monthly: [36255, 9064, 33236, 36257, 27193, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [5409, 5410, 10818, 10818, 10819, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 27743, monthly: [0, 3759, 0, 0, 3690, null, null, null, null, null, null, null] }
      ] },
      { name: "宏恩龍安", channel: "地區醫院", system: "宏恩", city: "台中市南區", products: [
          { series: "Ritalin LA", yearTarget: 314905, monthly: [64001, 0, 0, 25600, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2971, 0, 0, 1486, 1486, null, null, null, null, null, null, null] }
      ] },
      { name: "台中中山附醫中興分院", channel: "地區醫院", system: "台中中山", city: "台中市南區", products: [
          { series: "Ritalin LA", yearTarget: 181725, monthly: [0, 0, 35358, 0, 35358, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [10571, 0, 2643, 5286, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "員榮醫院", channel: "地區醫院", system: "員榮", city: "彰化縣員林市", products: [
          { series: "Exelon", yearTarget: 134591, monthly: [18129, 6646, 0, 7251, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 50304, monthly: [3759, 0, 7603, 1272, 6194, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [8114, 0, 0, 0, 5409, null, null, null, null, null, null, null] }
      ] },
      { name: "員郭醫院", channel: "地區醫院", system: "", city: "彰化縣員林市", products: [
          { series: "Ritalin LA", yearTarget: 42403, monthly: [30286, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2971, 3714, 0, 2972, 1486, null, null, null, null, null, null, null] }
      ] },
      { name: "常春醫院", channel: "地區醫院", system: "宏仁", city: "彰化縣員林市", products: [
          { series: "Exelon Patch", yearTarget: 16569, monthly: [4519, 0, 0, 4519, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 0, monthly: [0, 0, 6960, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "黃煒勝", group: "Div1", customers: [
      { name: "馬偕台北", channel: "醫學中心", system: "馬偕", city: "台北市中山區", products: [
          { series: "Ritalin LA", yearTarget: 1520210, monthly: [111389, 119108, 119108, 237665, 198515, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 1351997, monthly: [225332, 0, 135200, 45066, 90134, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 1205815, monthly: [123624, 128203, 128203, 128203, 45787, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1003781, monthly: [110729, 121986, 0, 148275, 52429, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [73371, 0, 48914, 97828, 97828, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 0, monthly: [0, 0, 0, 94776, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [8045, 8045, 8045, 0, 8045, null, null, null, null, null, null, null] }
      ] },
      { name: "北醫", channel: "區域醫院", system: "北醫", city: "台北市信義區", products: [
          { series: "Exelon", yearTarget: 2294506, monthly: [179523, 143123, 202313, 214883, 215880, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1016925, monthly: [82692, 37531, 105272, 64984, 93798, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 718360, monthly: [129647, 0, 55562, 83344, 120387, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [76047, 2746, 25766, 36073, 30631, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [22857, 0, 11428, 14285, 14285, null, null, null, null, null, null, null] }
      ] },
      { name: "恩主公", channel: "區域醫院", system: "", city: "新北市三峽區", products: [
          { series: "Ritalin LA", yearTarget: 1128364, monthly: [151543, 174274, 151543, 151543, 151543, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 1095425, monthly: [83280, 83280, 66624, 119368, 77728, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 563268, monthly: [109115, 0, 29714, 35765, 124757, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [113122, 0, 69690, 58047, 92322, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [22800, 26600, 22800, 17480, 18392, null, null, null, null, null, null, null] }
      ] },
      { name: "馬偕淡水", channel: "醫學中心", system: "馬偕", city: "新北市淡水區", products: [
          { series: "Exelon", yearTarget: 668230, monthly: [75547, 41209, 34340, 66390, 61812, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 616041, monthly: [0, 142163, 50746, 0, 189550, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 540799, monthly: [45066, 45066, 45066, 45066, 45066, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 517320, monthly: [50643, 38665, 46140, 50444, 42680, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 466484, monthly: [39703, 39703, 13235, 27572, 106977, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [85914, 0, 36814, 34400, 51600, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [12229, 12229, 12229, 27515, 15285, null, null, null, null, null, null, null] }
      ] },
      { name: "聖保祿", channel: "區域醫院", system: "長庚", city: "桃園市桃園區", products: [
          { series: "Exelon", yearTarget: 1913428, monthly: [239936, 94553, 130644, 198255, 147419, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 614825, monthly: [70400, 0, 42900, 51903, 20126, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 305100, monthly: [59814, 0, 17090, 17090, 17090, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [12124, 6411, 7419, 8505, 10391, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [13300, 0, 5666, 4149, 4602, null, null, null, null, null, null, null] }
      ] },
      { name: "台北台安", channel: "區域醫院", system: "", city: "台北市松山區", products: [
          { series: "Ritalin LA", yearTarget: 1270373, monthly: [149256, 22388, 111943, 149256, 74628, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 727731, monthly: [40320, 40320, 53760, 94080, 40320, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 499070, monthly: [43268, 21635, 57688, 50401, 50400, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [28572, 21429, 21429, 21429, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [9071, 9072, 3024, 7559, 11490, null, null, null, null, null, null, null] }
      ] },
      { name: "部八里療", channel: "區域醫院", system: "部標", city: "新北市八里區", products: [
          { series: "Ritalin LA", yearTarget: 1180750, monthly: [123229, 104163, 66216, 81356, 97301, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [26357, 7124, 17809, 15672, 17096, null, null, null, null, null, null, null] }
      ] },
      { name: "部樂生療", channel: "地區醫院", system: "部標", city: "桃園市龜山區", products: [
          { series: "Exelon", yearTarget: 557414, monthly: [43335, 28889, 79447, 0, 72224, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 63848, monthly: [5804, 7255, 4352, 4030, 15368, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4275, 2850, 2849, 9261, 3562, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [3825, 0, 3825, 2295, 9945, null, null, null, null, null, null, null] }
      ] },
      { name: "綠杏大藥局(北醫)", channel: "健保藥局", system: "1B", city: "台北市信義區", products: [
          { series: "Ritalin LA", yearTarget: 323086, monthly: [31095, 17986, 16462, 9857, 13245, null, null, null, null, null, null, null] }
      ] },
      { name: "博仁醫院", channel: "地區醫院", system: "", city: "台北市松山區", products: [
          { series: "Exelon", yearTarget: 154312, monthly: [12467, 4400, 4400, 14668, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市林森(慢性病)", channel: "區域醫院", system: "北市", city: "台北市中山區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [2857, 7143, 0, 7143, 3571, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 20006, monthly: [0, 0, 7219, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 1851, 0, 926, 1234, null, null, null, null, null, null, null] }
      ] },
      { name: "北市昆明", channel: "區域醫院", system: "北市", city: "台北市萬華區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 0, 0, 7143, 7143, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "許柔柔", group: "Div1", customers: [
      { name: "慈濟花蓮", channel: "醫學中心", system: "慈濟", city: "花蓮縣花蓮市", products: [
          { series: "Exelon", yearTarget: 2448391, monthly: [380320, 126773, 269527, 222120, 206140, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 1201968, monthly: [125743, 56429, 56429, 109315, 67726, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 899184, monthly: [60000, 19371, 60000, -1860, 60000, null, null, null, null, null, null, null] },
          { series: "Gilenya", yearTarget: 568210, monthly: [47350, 47350, 94702, 47350, 47350, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [58476, 0, 35086, 35086, 39306, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [20743, 0, 7779, 0, 18150, null, null, null, null, null, null, null] }
      ] },
      { name: "門諾醫院", channel: "區域醫院", system: "門諾", city: "花蓮縣花蓮市", products: [
          { series: "Gilenya", yearTarget: 1847544, monthly: [142118, 142118, 94746, 189490, 189490, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1215455, monthly: [93020, 93020, 93020, 128608, 93020, null, null, null, null, null, null, null] },
          { series: "Mayzent", yearTarget: 631929, monthly: [48609, 0, 97220, 0, 145830, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 621644, monthly: [42000, 42000, 42000, 0, 41503, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [19314, 9657, 9657, 28971, 14487, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [11619, 11619, 0, 16266, 9295, null, null, null, null, null, null, null] }
      ] },
      { name: "陽大附醫(蘭陽)", channel: "區域醫院", system: "陽大", city: "宜蘭縣宜蘭市", products: [
          { series: "Exelon", yearTarget: 1461953, monthly: [185748, 26421, 79423, 129283, 188730, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 438303, monthly: [94492, 21566, 14897, 55904, 7050, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [24229, 12114, 0, 12114, 24228, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 48274, monthly: [49508, 0, 24754, 55697, 24754, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [6500, 6500, 6500, 6500, 6500, null, null, null, null, null, null, null] }
      ] },
      { name: "羅東聖母", channel: "區域醫院", system: "聖母", city: "宜蘭縣羅東鎮", products: [
          { series: "Exelon", yearTarget: 909547, monthly: [74367, 61973, 61973, 130143, 74367, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 390573, monthly: [69857, 0, 69857, 0, 69857, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 147760, monthly: [22395, 6177, 13642, 28187, 14746, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [37143, 0, 0, 19782, 18571, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14858, 14858, 7429, 11143, 22286, null, null, null, null, null, null, null] }
      ] },
      { name: "陽大附醫(新民)", channel: "區域醫院", system: "陽大", city: "宜蘭縣宜蘭市", products: [
          { series: "Exelon Patch", yearTarget: 817098, monthly: [189503, 0, 0, 97960, 28731, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 263047, monthly: [59607, 39632, 59768, 13210, 26421, null, null, null, null, null, null, null] }
      ] },
      { name: "部玉里", channel: "地區醫院", system: "部標", city: "花蓮縣玉里鎮", products: [
          { series: "Ritalin LA", yearTarget: 806174, monthly: [85973, 76517, 65941, 75397, 107809, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [18538, 6179, 15449, 30896, 12360, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7124, 7124, 7124, 7124, 7124, null, null, null, null, null, null, null] }
      ] },
      { name: "神農養生藥局(開蘭安心診所)", channel: "診所(含門前藥局)", system: "", city: "宜蘭縣宜蘭市", products: [
          { series: "Exelon", yearTarget: 449261, monthly: [9800, 19600, 39200, 0, 75133, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 402970, monthly: [92571, 24000, 0, 0, 23143, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [9071, 0, 0, 13608, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1491, 1491, 0, 0, 1491, null, null, null, null, null, null, null] }
      ] },
      { name: "羅東博愛", channel: "區域醫院", system: "羅東博愛", city: "宜蘭縣羅東鎮", products: [
          { series: "Ritalin LA", yearTarget: 432471, monthly: [22835, 50743, 26640, 19007, 34885, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 191357, monthly: [20745, 19449, 0, 12169, 12658, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [23466, 47821, 2343, 24591, 27425, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [12683, 26286, 0, 13143, 19713, null, null, null, null, null, null, null] }
      ] },
      { name: "軍花蓮805", channel: "區域醫院", system: "軍標", city: "花蓮縣新城鄉", products: [
          { series: "Exelon", yearTarget: 263246, monthly: [28174, 23116, 20228, 16615, 18783, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 243506, monthly: [27110, 8601, 36568, 13984, 25683, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 188265, monthly: [21982, 14726, 14725, 12738, 17262, null, null, null, null, null, null, null] }
      ] },
      { name: "金英寶藥局(同心診所)", channel: "診所(含門前藥局)", system: "", city: "花蓮縣花蓮市", products: [
          { series: "Ritalin LA", yearTarget: 570546, monthly: [32643, 27477, 33095, 31193, 37588, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [37000, 22200, 22200, 29600, 5920, null, null, null, null, null, null, null] }
      ] },
      { name: "部花蓮", channel: "地區醫院", system: "部標", city: "花蓮縣花蓮市", products: [
          { series: "Ritalin LA", yearTarget: 517684, monthly: [52708, 14745, 63283, 12508, 49649, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7124, 7124, 14248, 14248, 14248, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 31601, monthly: [3010, 3010, 3010, 2903, 2903, null, null, null, null, null, null, null] }
      ] },
      { name: "慈濟玉里", channel: "地區醫院", system: "慈濟", city: "花蓮縣玉里鎮", products: [
          { series: "Exelon", yearTarget: 273294, monthly: [38032, 0, 25354, 18856, 9588, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 216772, monthly: [21528, 9704, 16880, 14351, 37023, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [9356, 0, 4678, 9356, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "玉里榮民", channel: "地區醫院", system: "北榮", city: "花蓮縣玉里鎮", products: [
          { series: "Exelon", yearTarget: 365981, monthly: [32027, 26881, 0, 51148, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 60358, monthly: [5600, 0, 0, 1041, 3107, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 23518, monthly: [0, 0, 6188, 0, 6188, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 14400, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1474, 0, 1474, 1474, null, null, null, null, null, null, null] }
      ] },
      { name: "蘇澳榮民", channel: "地區醫院", system: "北榮", city: "宜蘭縣蘇澳鎮", products: [
          { series: "Exelon Patch", yearTarget: 395064, monthly: [60199, 0, 5600, 34478, 26598, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3600, 0, 4319, 3600, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "員山榮民", channel: "地區醫院", system: "北榮", city: "宜蘭縣員山鄉", products: [
          { series: "Exelon Patch", yearTarget: 197276, monthly: [31880, 16200, 0, 21000, 15600, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 99643, monthly: [0, 0, 44557, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3600, 0, 0, 3600, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [737, 0, 0, 5406, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "賦康藥局(怡然診所)", channel: "診所(含門前藥局)", system: "羅東博愛", city: "宜蘭縣宜蘭市", products: [
          { series: "Ritalin LA", yearTarget: 105252, monthly: [0, 12454, 21241, 50543, 77769, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [6571, 0, 3943, 3286, 1971, null, null, null, null, null, null, null] }
      ] },
      { name: "雅信診所", channel: "診所(含門前藥局)", system: "", city: "宜蘭縣宜蘭市", products: [
          { series: "Ritalin LA", yearTarget: 35957, monthly: [16342, 0, 0, 16342, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4560, 0, 0, 1520, 3192, null, null, null, null, null, null, null] }
      ] },
      { name: "宇康藥局(悅增身心診所)", channel: "診所(含門前藥局)", system: "", city: "花蓮縣花蓮市", products: [
          { series: "Exelon Patch", yearTarget: 33944, monthly: [0, 0, 7714, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 5000, 5714, 2857, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "宜蘭員山", channel: "地區醫院", system: "蘭陽仁愛", city: "宜蘭縣宜蘭市", products: [
          { series: "Ritalin LA", yearTarget: 30691, monthly: [0, 0, 4286, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 2432, 2432, 2280, null, null, null, null, null, null, null] }
      ] },
      { name: "杏和醫院宜蘭", channel: "地區醫院", system: "聖母", city: "宜蘭縣礁溪鄉", products: [
          { series: "Exelon Patch", yearTarget: 14931, monthly: [4480, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 1238, 619, 2476, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 699, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [297, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安心藥局(平和身心診所)", channel: "診所(含門前藥局)", system: "", city: "宜蘭縣五結鄉", products: [
          { series: "Ritalin LA", yearTarget: 17740, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "鳳林榮民", channel: "地區醫院", system: "北榮", city: "花蓮縣鳳林鎮", products: [
          { series: "Exelon Patch", yearTarget: 11880, monthly: [0, 0, 0, 0, 1041, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [491, 1966, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "宜蘭仁愛", channel: "地區醫院", system: "蘭陽仁愛", city: "宜蘭縣宜蘭市", products: [
          { series: "Exelon Patch", yearTarget: 8568, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 2432, 0, 0, 1520, null, null, null, null, null, null, null] }
      ] },
      { name: "中山身心診所", channel: "診所(含門前藥局)", system: "", city: "花蓮縣花蓮市", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 3614, 1808, 3977, null, null, null, null, null, null, null] }
      ] },
      { name: "海天醫院", channel: "地區醫院", system: "", city: "宜蘭縣壯圍鄉", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [723, 0, 2169, 1446, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "OA空區", group: "Div1", customers: [
      { name: "亞東醫院", channel: "醫學中心", system: "亞東", city: "新北市板橋區", products: [
          { series: "Ritalin LA", yearTarget: 3282482, monthly: [293265, 195510, 306842, 293265, 389662, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 611501, monthly: [270000, 264000, 294000, 252000, 288000, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 481810, monthly: [42047, 42046, 42046, 27391, 40840, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [84686, 84687, 84687, 98801, 112915, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [50885, 16833, 67718, 40400, 28809, null, null, null, null, null, null, null] }
      ] },
      { name: "基隆長庚", channel: "區域醫院", system: "長庚", city: "基隆市安樂區", products: [
          { series: "Ritalin LA", yearTarget: 1462669, monthly: [122842, 83008, 139071, 170899, 27895, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 246547, monthly: [130643, 0, 74218, 85402, 77776, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 199717, monthly: [22320, 9120, 10260, 23585, 20486, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [54766, 1881, 37162, 32776, 18617, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7235, 7289, 8805, 7688, 7554, null, null, null, null, null, null, null] }
      ] },
      { name: "北市松德", channel: "區域醫院", system: "北市", city: "台北市信義區", products: [
          { series: "Ritalin LA", yearTarget: 1233548, monthly: [161259, 107507, 107507, 107507, 107507, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 453589, monthly: [44583, 66337, 36794, 51546, 37049, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 130360, monthly: [36093, 43313, 39703, 54140, 39703, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [17857, 14286, 21429, 35715, 28572, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [10800, 7714, 6172, 9257, 9257, null, null, null, null, null, null, null] }
      ] },
      { name: "暘基醫院", channel: "地區醫院", system: "暘基", city: "基隆市安樂區", products: [
          { series: "Ritalin LA", yearTarget: 1547468, monthly: [256820, 45823, 76765, 181843, 215195, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14286, 0, 14286, 7142, 3571, null, null, null, null, null, null, null] }
      ] },
      { name: "軍北投818", channel: "區域醫院", system: "三總", city: "台北市北投區", products: [
          { series: "Ritalin LA", yearTarget: 1114330, monthly: [176369, 0, 79966, 128893, 69697, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 123441, monthly: [108358, 0, 32508, 51292, 49124, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [35500, 0, 0, 41500, 16358, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 45144, monthly: [8707, 2902, 7256, 9782, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [10200, 0, 0, 6490, 2472, null, null, null, null, null, null, null] }
      ] },
      { name: "北市中興", channel: "區域醫院", system: "北市", city: "台北市大同區", products: [
          { series: "Ritalin LA", yearTarget: 902760, monthly: [0, 53753, 53753, 53753, 53753, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 140452, monthly: [51975, 38261, 29597, 46921, 43311, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14286, 14286, 21429, 21429, 28571, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 58008, monthly: [0, 5802, 2901, 2899, 7248, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3703, 4320, 3085, 3085, 6170, null, null, null, null, null, null, null] }
      ] },
      { name: "北市仁愛", channel: "區域醫院", system: "北市", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 630215, monthly: [55993, 37329, 37329, 37329, 37329, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 146953, monthly: [43311, 50530, 43312, 14438, 79405, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 75195, monthly: [16544, 4512, 7520, 4511, 4511, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [10801, 9257, 10801, 15429, 9257, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 7143, 7143, 7143, 21429, null, null, null, null, null, null, null] }
      ] },
      { name: "部基隆", channel: "區域醫院", system: "部標", city: "基隆市信義區", products: [
          { series: "Ritalin LA", yearTarget: 325890, monthly: [61213, 0, 0, 22955, 19129, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 238059, monthly: [97505, 72224, 86672, 158896, 46946, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 75240, monthly: [0, 0, 7524, 0, 7255, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5700, 7124, 7124, 7124, 14247, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [6178, 4943, 0, 10505, 3708, null, null, null, null, null, null, null] }
      ] },
      { name: "北市忠孝", channel: "區域醫院", system: "北市", city: "台北市南港區", products: [
          { series: "Exelon Patch", yearTarget: 355083, monthly: [29543, 29543, 44046, 29532, 14497, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 92382, monthly: [28874, 36093, 57748, 28874, 43312, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 88111, monthly: [0, 0, 15678, 0, 7466, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [12344, 7406, 6172, 6172, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市和平", channel: "區域醫院", system: "北市", city: "台北市中正區", products: [
          { series: "Exelon Patch", yearTarget: 223384, monthly: [8702, 17726, 26427, 7249, 22015, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 103791, monthly: [0, 0, 7466, 18665, 14931, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 92722, monthly: [43314, 28876, 50533, 57751, 28876, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [7714, 4629, 7714, 9255, 6170, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 4286, 2857, 5714, 2143, null, null, null, null, null, null, null] }
      ] },
      { name: "遠東聯合診所(亞東)", channel: "診所(含門前藥局)", system: "亞東", city: "台北市中正區", products: [
          { series: "Exelon Patch", yearTarget: 158839, monthly: [29219, 18172, 9798, 19035, 13494, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 124291, monthly: [59399, 39600, 49800, 19200, 31800, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3529, 3529, 7057, 3529, 3529, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [2886, 1924, 962, 481, 1443, null, null, null, null, null, null, null] }
      ] },
      { name: "長庚情人湖", channel: "區域醫院", system: "長庚", city: "基隆市安樂區", products: [
          { series: "Ritalin LA", yearTarget: 277348, monthly: [15870, 4920, 33980, 25395, 23823, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 37699, monthly: [5591, 10167, 11183, 17793, 13218, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [5830, 0, 2172, 6153, 269, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 11058, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [745, 1064, 798, 798, 1729, null, null, null, null, null, null, null] }
      ] },
      { name: "123宅藥局", channel: "診所(含門前藥局)", system: "1B", city: "台北市大同區", products: [
          { series: "Ritalin LA", yearTarget: 293302, monthly: [32894, 14625, 17027, 18323, 27007, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1486, 0, 1486, 1114, null, null, null, null, null, null, null] }
      ] },
      { name: "軍松山807", channel: "區域醫院", system: "三總", city: "台北市松山區", products: [
          { series: "Exelon Patch", yearTarget: 144461, monthly: [13060, 3009, 22035, 6019, 7149, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 37955, monthly: [17041, 0, 0, 0, 3098, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1000, -786, 2857, 0, 1000, null, null, null, null, null, null, null] }
      ] },
      { name: "新北市三重", channel: "區域醫院", system: "部標", city: "新北市三重區", products: [
          { series: "Exelon Patch", yearTarget: 138442, monthly: [0, 10534, 9029, 17412, 15962, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 14247, 0, 11398, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "宏恩台北", channel: "地區醫院", system: "", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 95741, monthly: [15540, 0, 0, 35992, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1486, 0, 0, 1114, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "維德醫院", channel: "地區醫院", system: "", city: "基隆市中正區", products: [
          { series: "Exelon", yearTarget: 31962, monthly: [3587, 21520, 7173, 7173, 21520, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 21819, monthly: [3062, 0, 3062, 1486, 4458, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2971, 1486, 1486, 1486, 1486, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [963, 1605, 1926, 1605, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市婦幼", channel: "區域醫院", system: "北市", city: "台北市中正區", products: [
          { series: "Ritalin LA", yearTarget: 20908, monthly: [0, 7466, 7466, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3085, 3085, 1234, 5554, 1851, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 714, 0, 714, 714, null, null, null, null, null, null, null] }
      ] },
      { name: "新北市板橋", channel: "區域醫院", system: "部標", city: "新北市板橋區", products: [
          { series: "Exelon Patch", yearTarget: 25582, monthly: [0, 0, 0, 0, 2904, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1425, 0, 0, 1425, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "康科特-郵政醫院", channel: "地區醫院", system: "康科特", city: "台北市中正區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [4731, 1183, 5914, 4731, 2957, null, null, null, null, null, null, null] }
      ] },
      { name: "和信治癌", channel: "區域醫院", system: "", city: "台北市北投區", products: [
          { series: "Exelon", yearTarget: 4376, monthly: [0, 0, 3693, 3694, 3693, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, -1308, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "軍基隆812", channel: "地區醫院", system: "三總", city: "基隆市中正區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 2472, 3709, 3709, 1854, null, null, null, null, null, null, null] }
      ] },
      { name: "北市忠孝(信義門診)", channel: "區域醫院", system: "北市", city: "台北市信義區", products: [
          { series: "Exelon", yearTarget: 0, monthly: [3610, 7219, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "礦工醫院", channel: "地區醫院", system: "礦工", city: "基隆市暖暖區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [2762, 0, 0, 4143, 1381, null, null, null, null, null, null, null] }
      ] },
      { name: "台北中山醫院", channel: "地區醫院", system: "", city: "台北市大安區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [1800, 0, 2099, 1500, 600, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [380, 760, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "侯靜琦", group: "Div1", customers: [
      { name: "亞東醫院", channel: "醫學中心", system: "亞東", city: "新北市板橋區", products: [
          { series: "Exelon", yearTarget: 2036985, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "基隆長庚", channel: "區域醫院", system: "長庚", city: "基隆市安樂區", products: [
          { series: "Exelon", yearTarget: 821279, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部基隆", channel: "區域醫院", system: "部標", city: "基隆市信義區", products: [
          { series: "Exelon", yearTarget: 793005, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市仁愛", channel: "區域醫院", system: "北市", city: "台北市大安區", products: [
          { series: "Exelon", yearTarget: 489527, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市中興", channel: "區域醫院", system: "北市", city: "台北市大同區", products: [
          { series: "Exelon", yearTarget: 467862, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市松德", channel: "區域醫院", system: "北市", city: "台北市信義區", products: [
          { series: "Exelon", yearTarget: 434246, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "遠東聯合診所(亞東)", channel: "診所(含門前藥局)", system: "亞東", city: "台北市中正區", products: [
          { series: "Exelon", yearTarget: 414028, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "軍北投818", channel: "區域醫院", system: "三總", city: "台北市北投區", products: [
          { series: "Exelon", yearTarget: 411194, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市和平", channel: "區域醫院", system: "北市", city: "台北市中正區", products: [
          { series: "Exelon", yearTarget: 308875, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市忠孝", channel: "區域醫院", system: "北市", city: "台北市南港區", products: [
          { series: "Exelon", yearTarget: 307733, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "長庚情人湖", channel: "區域醫院", system: "長庚", city: "基隆市安樂區", products: [
          { series: "Exelon", yearTarget: 125574, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "維德醫院", channel: "地區醫院", system: "", city: "基隆市中正區", products: [
          { series: "Exelon", yearTarget: 106472, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "和信治癌", channel: "區域醫院", system: "", city: "台北市北投區", products: [
          { series: "Exelon", yearTarget: 14579, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "王維逸", group: "Div2", customers: [
      { name: "北市忠孝", channel: "區域醫院", system: "北市", city: "台北市南港區", products: [
          { series: "Lote", yearTarget: 1366750, monthly: [155676, 99067, 99067, 99067, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [4800, 4800, 4800, 4800, 4800, null, null, null, null, null, null, null] }
      ] },
      { name: "軍松山807", channel: "區域醫院", system: "三總", city: "台北市松山區", products: [
          { series: "Pane", yearTarget: 732044, monthly: [61200, 100800, 86400, 146160, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [1287, 0, 7404, 0, 4185, null, null, null, null, null, null, null] }
      ] },
      { name: "軍三總", channel: "醫學中心", system: "三總", city: "台北市內湖區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [115886, 186857, 0, 0, 186857, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [20000, 5000, 8333, 13333, 10000, null, null, null, null, null, null, null] }
      ] },
      { name: "軍新竹813", channel: "地區醫院", system: "軍標", city: "新竹市北區", products: [
          { series: "Ezole", yearTarget: 105300, monthly: [52657, 96538, 52657, 35067, 61367, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [16095, 19314, 12876, 0, 12876, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 31830, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 28747, monthly: [2865, 2865, 2865, 5730, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部桃療", channel: "區域醫院", system: "部標", city: "桃園市桃園區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [268473, 0, 9588, 90130, 100358, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [2421, 3767, 4304, 5650, 2420, null, null, null, null, null, null, null] }
      ] },
      { name: "維德醫院", channel: "地區醫院", system: "", city: "基隆市中正區", products: [
          { series: "Epine", yearTarget: 124926, monthly: [18906, 10605, 10580, 8839, 8586, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 117780, monthly: [13047, 19381, 9094, 0, 9095, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [12043, 17076, 10164, 7927, 7622, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 24467, monthly: [2843, 4549, 6254, 1643, 1095, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 15870, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "軍北投818", channel: "區域醫院", system: "三總", city: "台北市北投區", products: [
          { series: "Lote", yearTarget: 226950, monthly: [44403, 0, 0, 4297, 12891, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 45962, monthly: [2627, 0, 0, 7816, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 37950, monthly: [15940, 0, 0, 7590, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 19440, monthly: [0, 0, 0, 1438, 2876, null, null, null, null, null, null, null] }
      ] },
      { name: "東元綜合醫院", channel: "區域醫院", system: "", city: "新竹縣竹北市", products: [
          { series: "Lote", yearTarget: 157950, monthly: [36457, 0, 0, 14583, 14583, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 29841, monthly: [0, 0, 0, 2400, 2057, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [13714, 0, 0, 6857, 6857, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 24320, monthly: [11429, 0, 5333, 7619, 0, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 5670, monthly: [24286, 0, 12143, 13762, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部樂生療", channel: "地區醫院", system: "部標", city: "桃園市龜山區", products: [
          { series: "Pane", yearTarget: 259521, monthly: [15534, 20713, 20712, 15534, 10356, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [5388, 0, 0, 2694, 1347, null, null, null, null, null, null, null] }
      ] },
      { name: "部台北", channel: "區域醫院", system: "部標", city: "新北市新莊區", products: [
          { series: "Ezole", yearTarget: 99750, monthly: [34171, 39866, 17086, 23700, 18433, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [40399, 0, 14813, 24239, 18853, null, null, null, null, null, null, null] }
      ] },
      { name: "亞東醫院", channel: "醫學中心", system: "亞東", city: "新北市板橋區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [45000, 30000, 30000, 30000, 30000, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [22286, 44572, 22286, 44572, 22286, null, null, null, null, null, null, null] }
      ] },
      { name: "輔大附醫", channel: "區域醫院", system: "", city: "新北市泰山區", products: [
          { series: "Epram", yearTarget: 223100, monthly: [18476, 9238, 18476, 18476, 18476, null, null, null, null, null, null, null] }
      ] },
      { name: "台北台安", channel: "區域醫院", system: "", city: "台北市松山區", products: [
          { series: "Epram", yearTarget: 177100, monthly: [16429, 10952, 18619, 14238, 21904, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [4771, 4771, 7953, 12724, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "景美醫院", channel: "地區醫院", system: "佳醫集康", city: "台北市文山區", products: [
          { series: "Lote", yearTarget: 49962, monthly: [6058, 4543, 0, 7571, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 32760, monthly: [4095, 0, 4095, 3872, 0, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 31535, monthly: [7048, 1762, 0, 3524, 3524, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 14619, monthly: [3648, 0, 0, 3648, 3648, null, null, null, null, null, null, null] }
      ] },
      { name: "中國新竹", channel: "地區醫院", system: "中國", city: "新竹縣竹北市", products: [
          { series: "Lote", yearTarget: 74250, monthly: [4952, 7924, 7924, 9905, 12876, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 25200, monthly: [1048, 4190, 0, 2095, 6285, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [5638, 0, 2819, 4229, 3524, null, null, null, null, null, null, null] }
      ] },
      { name: "台大新竹", channel: "醫學中心", system: "台大", city: "新竹市北區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [54364, 15134, 28180, 41285, 30266, null, null, null, null, null, null, null] }
      ] },
      { name: "台大新竹生醫", channel: "區域醫院", system: "台大", city: "新竹縣竹北市", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [31619, 31619, 11048, 31619, 31619, null, null, null, null, null, null, null] }
      ] },
      { name: "軍桃園804", channel: "區域醫院", system: "軍標", city: "桃園市龍潭區", products: [
          { series: "Fute", yearTarget: 47745, monthly: [5305, 0, 0, 5305, 10610, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 40851, monthly: [4297, 0, 5729, 4297, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 12285, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新北市三重", channel: "區域醫院", system: "部標", city: "新北市三重區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [44746, 0, 0, 0, 38353, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [10773, 8080, 8080, 9426, 9426, null, null, null, null, null, null, null] }
      ] },
      { name: "軍基隆812", channel: "地區醫院", system: "三總", city: "基隆市中正區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [19314, 19314, 19314, 19314, 19314, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [3333, 0, 3333, 3333, 1667, null, null, null, null, null, null, null] }
      ] },
      { name: "北市忠孝(信義門診)", channel: "區域醫院", system: "北市", city: "台北市信義區", products: [
          { series: "Lote", yearTarget: 49525, monthly: [21228, 14152, 0, 21229, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市仁愛", channel: "區域醫院", system: "北市", city: "台北市大安區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [6400, 4800, 4800, 4800, 6400, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 20907, monthly: [2726, 0, 2726, 0, 2720, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 0, 16429, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [1695, 1695, 3391, 3391, 3391, null, null, null, null, null, null, null] }
      ] },
      { name: "部桃園", channel: "區域醫院", system: "部標", city: "桃園市桃園區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [21546, 13466, 13466, 13466, 20199, null, null, null, null, null, null, null] }
      ] },
      { name: "天晟醫院(天成生醫)", channel: "區域醫院", system: "天成", city: "桃園市中壢區", products: [
          { series: "Bisadyl Supp", yearTarget: 38295, monthly: [11666, 0, 0, 0, 2667, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [3810, 0, 0, 5714, 1905, null, null, null, null, null, null, null] }
      ] },
      { name: "桃園榮民", channel: "區域醫院", system: "北榮", city: "桃園市桃園區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [11571, 0, 12729, 4628, 4628, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 1612, 7253, 4028, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 6411, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "天祥醫院(天成生醫)", channel: "地區醫院", system: "天成", city: "桃園市中壢區", products: [
          { series: "Tone", yearTarget: 13940, monthly: [8456, 0, 0, 3276, 4228, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 10989, monthly: [6667, 0, 0, 1667, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部桃園(新屋)", channel: "地區醫院", system: "部標", city: "桃園市新屋區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [9426, 0, 9426, 0, 16160, null, null, null, null, null, null, null] }
      ] },
      { name: "北醫萬芳", channel: "醫學中心", system: "北醫", city: "台北市文山區", products: [
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 0, 28000, null, null, null, null, null, null, null] }
      ] },
      { name: "天成醫院(天成生醫)", channel: "地區醫院", system: "天成", city: "桃園市楊梅區", products: [
          { series: "Bisadyl Supp", yearTarget: 15984, monthly: [0, 1333, 0, 0, 3333, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [1905, 952, 1905, 0, 1905, null, null, null, null, null, null, null] }
      ] },
      { name: "部基隆", channel: "區域醫院", system: "部標", city: "基隆市信義區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [5387, 9426, 0, 5386, 4040, null, null, null, null, null, null, null] }
      ] },
      { name: "北市陽明", channel: "區域醫院", system: "北市", city: "台北市士林區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 13562, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [4800, 3200, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "聯新-聯新", channel: "區域醫院", system: "聯新", city: "桃園市平鎮區", products: [
          { series: "Fute", yearTarget: 16340, monthly: [1714, 0, 1714, 857, 857, null, null, null, null, null, null, null] }
      ] },
      { name: "新北市板橋", channel: "區域醫院", system: "部標", city: "新北市板橋區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 6393, 0, 0, 6393, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1347, 1347, 1347, 1347, 1347, null, null, null, null, null, null, null] }
      ] },
      { name: "台北中山醫院", channel: "地區醫院", system: "", city: "台北市大安區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [3657, 914, 2286, 3200, 5447, null, null, null, null, null, null, null] }
      ] },
      { name: "新北台安", channel: "地區醫院", system: "", city: "新北市三芝區", products: [
          { series: "Fute", yearTarget: 10000, monthly: [0, 0, 0, 0, 5000, null, null, null, null, null, null, null] }
      ] },
      { name: "新竹榮民", channel: "地區醫院", system: "北榮", city: "新竹縣竹東鎮", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 4628, 4628, 4628, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "聯新-平鎮延平藥局", channel: "診所(含門前藥局)", system: "聯新", city: "桃園市平鎮區", products: [
          { series: "Fute", yearTarget: 0, monthly: [0, 857, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "胡辰翰", group: "Div2", customers: [
      { name: "北市松德", channel: "區域醫院", system: "北市", city: "台北市信義區", products: [
          { series: "Lote", yearTarget: 937200, monthly: [70762, 66516, 76422, 79253, 77838, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 827480, monthly: [0, 66419, 66419, 24152, 84533, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [81371, 40686, 81371, 81371, 40686, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 212520, monthly: [26476, 39714, 39714, 26476, 26476, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 192708, monthly: [13629, 9994, 15446, 12693, 12693, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 171600, monthly: [15333, 12048, 17524, 9857, 30667, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 141351, monthly: [13371, 12257, 9286, 6686, 5200, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [4800, 4800, 6400, 6400, 4800, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 2800, monthly: [1410, 0, 0, 0, 705, null, null, null, null, null, null, null] }
      ] },
      { name: "海天醫院", channel: "地區醫院", system: "", city: "宜蘭縣壯圍鄉", products: [
          { series: "Exprexa", yearTarget: 793699, monthly: [55467, 0, 55467, 41333, 41333, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 310338, monthly: [31238, 37619, 26762, 16857, 30381, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 136655, monthly: [0, 18095, 0, 17714, 17714, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [10657, 0, 10657, 0, 10657, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1240, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "宜蘭仁愛", channel: "地區醫院", system: "蘭陽仁愛", city: "宜蘭縣宜蘭市", products: [
          { series: "Tonex", yearTarget: 90510, monthly: [8229, 8229, 8229, 8229, 8229, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 53795, monthly: [5571, 11143, 3714, 7429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北醫雙和", channel: "醫學中心", system: "北醫", city: "新北市中和區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [33524, 33524, 41904, 33524, 33524, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 0, 18667, null, null, null, null, null, null, null] }
      ] },
      { name: "宜蘭員山", channel: "地區醫院", system: "蘭陽仁愛", city: "宜蘭縣宜蘭市", products: [
          { series: "Bisadyl Supp", yearTarget: 123543, monthly: [11143, 11143, 0, 22286, 7429, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 12342, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "蘇澳榮民", channel: "地區醫院", system: "北榮", city: "宜蘭縣蘇澳鎮", products: [
          { series: "Bisadyl Supp", yearTarget: 89304, monthly: [12168, 14601, 0, 0, 14601, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 6943, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "黎明健康藥局(光中身心診所)", channel: "診所(含門前藥局)", system: "", city: "宜蘭縣宜蘭市", products: [
          { series: "Epine", yearTarget: 59275, monthly: [0, 19047, 0, 15714, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 0, 12001, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北醫", channel: "區域醫院", system: "北醫", city: "台北市信義區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [23468, 6705, 30172, 20114, 20116, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 0, 4667, null, null, null, null, null, null, null] }
      ] },
      { name: "員山榮民", channel: "地區醫院", system: "北榮", city: "宜蘭縣員山鄉", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [13886, 0, 20829, 16199, 0, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 33916, monthly: [6083, 0, 0, 3650, 3650, null, null, null, null, null, null, null] }
      ] },
      { name: "神農養生藥局(開蘭安心診所)", channel: "診所(含門前藥局)", system: "", city: "宜蘭縣宜蘭市", products: [
          { series: "Suculin", yearTarget: 52464, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [1786, 0, 0, 0, 1786, null, null, null, null, null, null, null] }
      ] },
      { name: "北市和平", channel: "區域醫院", system: "北市", city: "台北市中正區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [11200, 6400, 12800, 6400, 12800, null, null, null, null, null, null, null] }
      ] },
      { name: "北市中興", channel: "區域醫院", system: "北市", city: "台北市大同區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [3200, 3200, 4800, 3200, 3200, null, null, null, null, null, null, null] }
      ] },
      { name: "羅東聖母", channel: "區域醫院", system: "聖母", city: "宜蘭縣羅東鎮", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 5143, 0, 5143, 5143, null, null, null, null, null, null, null] }
      ] },
      { name: "中心綜合醫院", channel: "地區醫院", system: "", city: "台北市大安區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [3238, 0, 4857, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "杏和醫院宜蘭", channel: "地區醫院", system: "聖母", city: "宜蘭縣礁溪鄉", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [1715, 0, 1715, 1715, 1715, null, null, null, null, null, null, null] }
      ] },
      { name: "安心藥局(平和身心診所)", channel: "診所(含門前藥局)", system: "", city: "宜蘭縣五結鄉", products: [
          { series: "Ezole", yearTarget: 0, monthly: [2400, 400, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 0, monthly: [0, 2285, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "忠欣骨科診所", channel: "診所(含門前藥局)", system: "", city: "宜蘭縣羅東鎮", products: [
          { series: "Tone", yearTarget: 4275, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市婦幼", channel: "區域醫院", system: "北市", city: "台北市中正區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 1600, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "黃冠瑜", group: "Div2", customers: [
      { name: "部草屯療", channel: "區域醫院", system: "部標", city: "南投縣草屯鎮", products: [
          { series: "Fute", yearTarget: 434920, monthly: [31428, 31428, 0, 31428, 31428, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 219272, monthly: [20381, 20381, 0, 17089, 20057, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 145350, monthly: [28476, 31893, 0, 27405, 29493, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [13424, 13743, 12785, 12785, 13424, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [4040, 2693, 2693, 4040, 4040, null, null, null, null, null, null, null] }
      ] },
      { name: "部台中", channel: "區域醫院", system: "部標", city: "台中市西區", products: [
          { series: "Epram", yearTarget: 484380, monthly: [31000, 31000, 33067, 35134, 23767, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 246280, monthly: [15715, 15715, 15715, 31430, 15715, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [5386, 2693, 8080, 4040, 5386, null, null, null, null, null, null, null] }
      ] },
      { name: "中國附醫(急重症)", channel: "醫學中心", system: "中國", city: "台中市北區", products: [
          { series: "Lote", yearTarget: 488070, monthly: [77258, 29714, 54476, 39620, 59430, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 121180, monthly: [9428, 9428, 13618, 17285, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [16914, 4229, 18324, 12686, 12686, null, null, null, null, null, null, null] }
      ] },
      { name: "埔里基督教", channel: "區域醫院", system: "埔基", city: "南投縣埔里鎮", products: [
          { series: "Bisadyl Supp", yearTarget: 235300, monthly: [21714, 5429, 21714, 12667, 18095, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 44080, monthly: [2324, 0, 3486, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 31350, monthly: [2471, 824, 3295, 2471, 4119, null, null, null, null, null, null, null] }
      ] },
      { name: "員郭醫院", channel: "地區醫院", system: "", city: "彰化縣員林市", products: [
          { series: "Epine", yearTarget: 90605, monthly: [18056, 6171, 0, 10448, 9019, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 88827, monthly: [14130, 8074, 0, 10080, 10080, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 50160, monthly: [6600, 6600, 0, 7920, 3960, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 24684, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 11106, monthly: [2468, 0, 0, 2467, 3700, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 9947, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "彰化基督教", channel: "醫學中心", system: "彰基", city: "彰化縣員林市", products: [
          { series: "Lote", yearTarget: 239547, monthly: [26571, 26571, 35428, 35428, 26571, null, null, null, null, null, null, null] }
      ] },
      { name: "常春醫院", channel: "地區醫院", system: "宏仁", city: "彰化縣員林市", products: [
          { series: "Bisadyl Supp", yearTarget: 216660, monthly: [37714, 37714, 0, 37714, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部彰化", channel: "區域醫院", system: "部標", city: "彰化縣埔心鄉", products: [
          { series: "Fute", yearTarget: 220080, monthly: [10476, 20952, 26190, 5238, 41904, null, null, null, null, null, null, null] }
      ] },
      { name: "宏仁醫院", channel: "地區醫院", system: "宏仁", city: "彰化縣員林市", products: [
          { series: "Bisadyl Supp", yearTarget: 173956, monthly: [37714, 37714, 37714, 0, 37714, null, null, null, null, null, null, null] }
      ] },
      { name: "彰濱秀傳", channel: "區域醫院", system: "秀傳", city: "彰化縣鹿港鎮", products: [
          { series: "Fute", yearTarget: 171000, monthly: [42771, 23762, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "員榮醫院", channel: "地區醫院", system: "員榮", city: "彰化縣員林市", products: [
          { series: "Slatone", yearTarget: 59500, monthly: [8476, 4238, 4238, 4238, 4238, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 20720, monthly: [2952, 0, 0, 1476, 1476, null, null, null, null, null, null, null] }
      ] },
      { name: "若瑟醫院", channel: "地區醫院", system: "", city: "雲林縣虎尾鎮", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [45514, 10115, 16858, 16858, 18542, null, null, null, null, null, null, null] }
      ] },
      { name: "東華醫院", channel: "地區醫院", system: "宏仁", city: "南投縣竹山鎮", products: [
          { series: "Bisadyl Supp", yearTarget: 65312, monthly: [37714, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "埔里榮民", channel: "地區醫院", system: "中榮", city: "南投縣埔里鎮", products: [
          { series: "Fute", yearTarget: 61380, monthly: [13644, 0, 0, 6822, 6822, null, null, null, null, null, null, null] }
      ] },
      { name: "部南投", channel: "區域醫院", system: "部標", city: "南投縣南投市", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [24240, 8080, 16160, 16160, 8080, null, null, null, null, null, null, null] }
      ] },
      { name: "佑民醫院", channel: "地區醫院", system: "", city: "南投縣草屯鎮", products: [
          { series: "Lote", yearTarget: 59721, monthly: [0, 2074, 2074, 2074, 2074, null, null, null, null, null, null, null] }
      ] },
      { name: "員榮員生", channel: "地區醫院", system: "員榮", city: "彰化縣員林市", products: [
          { series: "Tone", yearTarget: 22200, monthly: [2952, 0, 2952, 2952, 0, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 22100, monthly: [3390, 0, 5086, 3390, 2543, null, null, null, null, null, null, null] }
      ] },
      { name: "彰化秀傳", channel: "區域醫院", system: "秀傳", city: "彰化縣彰化市", products: [
          { series: "Fute", yearTarget: 38000, monthly: [4752, 4752, 0, 4752, 4752, null, null, null, null, null, null, null] }
      ] },
      { name: "澄清中港", channel: "區域醫院", system: "", city: "台中市西屯區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [14000, 0, 7000, 8400, 8400, null, null, null, null, null, null, null] }
      ] },
      { name: "老人復健綜合醫院", channel: "地區醫院", system: "中國", city: "台中市北屯區", products: [
          { series: "Lote", yearTarget: 15840, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 9450, monthly: [0, 0, 0, 1048, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "員林郭醫院大村", channel: "地區醫院", system: "", city: "彰化縣大村鄉", products: [
          { series: "Tonex", yearTarget: 18446, monthly: [0, 0, 7378, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部豐原", channel: "區域醫院", system: "部標", city: "台中市豐原區", products: [
          { series: "Tone", yearTarget: 0, monthly: [4514, 6771, 0, 4518, 4518, null, null, null, null, null, null, null] }
      ] },
      { name: "林新醫院", channel: "區域醫院", system: "林新", city: "台中市南屯區", products: [
          { series: "Tone", yearTarget: 0, monthly: [2743, 0, 2743, 3657, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林新醫院烏日", channel: "地區醫院", system: "林新", city: "台中市烏日區", products: [
          { series: "Tone", yearTarget: 0, monthly: [2743, 2743, 0, 2743, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "竹山秀傳", channel: "地區醫院", system: "秀傳", city: "南投縣竹山鎮", products: [
          { series: "Fute", yearTarget: 4725, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "中國台中東區", channel: "地區醫院", system: "中國", city: "台中市東區", products: [
          { series: "Tone", yearTarget: 0, monthly: [705, 705, 0, 705, 705, null, null, null, null, null, null, null] }
      ] },
      { name: "亞洲附醫豐中分院", channel: "地區醫院", system: "中國", city: "台中市豐原區", products: [
          { series: "Tone", yearTarget: 0, monthly: [705, 0, 0, 705, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "惠和醫院", channel: "地區醫院", system: "中國", city: "南投縣草屯鎮", products: [
          { series: "Tone", yearTarget: 0, monthly: [705, 0, 705, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "惠盛醫院", channel: "地區醫院", system: "中國", city: "台中市豐原區", products: [
          { series: "Tone", yearTarget: 0, monthly: [1410, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "王立穎", group: "Div2", customers: [
      { name: "為恭醫院", channel: "區域醫院", system: "中國", city: "苗栗縣頭份市", products: [
          { series: "Lote", yearTarget: 742500, monthly: [128762, 0, 69334, 34667, 39620, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 309840, monthly: [28285, 13619, 24095, 13095, 27238, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [3524, 2114, 0, 3524, 1410, null, null, null, null, null, null, null] }
      ] },
      { name: "宏恩龍安", channel: "地區醫院", system: "宏恩", city: "台中市南區", products: [
          { series: "Exprexa", yearTarget: 539382, monthly: [71573, 0, 24448, 16267, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 170150, monthly: [16096, 0, 16096, 0, 3429, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 150380, monthly: [8800, 0, 4400, 16838, 2190, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 101150, monthly: [14286, 0, 0, 4762, 7143, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 70563, monthly: [4867, 0, 4867, 0, 9728, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [5981, 0, 5980, 0, 0, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 10017, monthly: [743, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 9564, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 4032, monthly: [4032, 0, 0, 0, 4032, null, null, null, null, null, null, null] },
          { series: "Cospirit", yearTarget: 1819, monthly: [1841, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "台中中山附醫", channel: "醫學中心", system: "台中中山", city: "台中市南區", products: [
          { series: "Epram", yearTarget: 205700, monthly: [95362, 0, 21771, 21771, 14514, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 151510, monthly: [20857, 6952, 13904, 20857, 13905, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 116748, monthly: [12686, 8457, 16914, 16914, 8457, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [4990, 0, 2495, 2495, 2495, null, null, null, null, null, null, null] }
      ] },
      { name: "童綜合醫院", channel: "區域醫院", system: "童綜合", city: "台中市梧棲區", products: [
          { series: "Lote", yearTarget: 329530, monthly: [53143, 31886, 31886, 31886, 31886, null, null, null, null, null, null, null] }
      ] },
      { name: "李綜合大甲", channel: "區域醫院", system: "李綜合", city: "台中市大甲區", products: [
          { series: "Epine", yearTarget: 99790, monthly: [7096, 10362, 7096, 12019, 6943, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 60185, monthly: [3124, 4476, 1562, 12076, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 35552, monthly: [0, 11114, 0, 8891, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 33869, monthly: [2507, 2507, 2507, 2126, 2347, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3124, 0, 0, 6248, 3124, null, null, null, null, null, null, null] }
      ] },
      { name: "李綜合苑裡", channel: "地區醫院", system: "李綜合", city: "苗栗縣苑裡鎮", products: [
          { series: "Epine", yearTarget: 56600, monthly: [10820, 0, 6429, 447, 6381, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 51106, monthly: [20005, 0, 0, 6669, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 49549, monthly: [5014, 3760, 6266, 1870, 5867, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 41835, monthly: [0, 8952, 0, 0, 4476, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 17145, monthly: [3810, 0, 1905, 0, 1905, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 1562, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "美德醫院", channel: "地區醫院", system: "李綜合", city: "台中市大甲區", products: [
          { series: "Exprexa", yearTarget: 49549, monthly: [7520, 6267, 0, 4233, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 35310, monthly: [4286, 4286, 4286, -260, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 28410, monthly: [0, 1562, 0, 0, 8952, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 27775, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 12495, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 762, 1434, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "台中榮總", channel: "醫學中心", system: "中榮", city: "台中市西屯區", products: [
          { series: "Fute", yearTarget: 148676, monthly: [0, 20466, 13644, 0, 13644, null, null, null, null, null, null, null] }
      ] },
      { name: "大千醫院", channel: "地區醫院", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Tone", yearTarget: 0, monthly: [30286, 0, 15143, 10095, 30286, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 33720, monthly: [0, 0, 0, 0, 13714, null, null, null, null, null, null, null] }
      ] },
      { name: "大千南勢", channel: "地區醫院", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Fute", yearTarget: 71092, monthly: [13714, 16000, 16000, 0, 2286, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [5048, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "宏仁診所", channel: "診所(含門前藥局)", system: "中國", city: "苗栗縣竹南鎮", products: [
          { series: "Lote", yearTarget: 69300, monthly: [19810, 0, 0, 0, 19810, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 2520, monthly: [2095, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [1410, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "弘大醫院", channel: "地區醫院", system: "", city: "苗栗縣苗栗市", products: [
          { series: "Tonex", yearTarget: 37699, monthly: [0, 4435, 4435, 4435, 4435, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 35232, monthly: [0, 0, 0, 2202, 2202, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 4163, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [1219, 0, 0, 1219, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "長安醫院", channel: "地區醫院", system: "", city: "台中市太平區", products: [
          { series: "Fute", yearTarget: 69660, monthly: [6429, 3857, 3857, 3857, 3857, null, null, null, null, null, null, null] }
      ] },
      { name: "光田向上", channel: "區域醫院", system: "光田", city: "台中市沙鹿區", products: [
          { series: "Fute", yearTarget: 52875, monthly: [-4672, 0, 3524, 7048, 7048, null, null, null, null, null, null, null] }
      ] },
      { name: "宏恩台中", channel: "地區醫院", system: "宏恩", city: "台中市南區", products: [
          { series: "Tonex", yearTarget: 20160, monthly: [4032, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3420, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部苗栗", channel: "區域醫院", system: "部標", city: "苗栗縣苗栗市", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [4040, 2693, 2693, 2693, 4040, null, null, null, null, null, null, null] }
      ] },
      { name: "慈祐醫院", channel: "地區醫院", system: "", city: "苗栗縣竹南鎮", products: [
          { series: "Bisadyl Supp", yearTarget: 4572, monthly: [0, 3810, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "李綜合順安", channel: "地區醫院", system: "李綜合", city: "台中市大甲區", products: [
          { series: "Epine", yearTarget: 4330, monthly: [0, 0, 0, 1033, 562, null, null, null, null, null, null, null] }
      ] },
      { name: "協和醫院", channel: "地區醫院", system: "", city: "苗栗縣苗栗市", products: [
          { series: "Tone", yearTarget: 1690, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大千大順", channel: "地區醫院", system: "大千", city: "苗栗縣苗栗市", products: [
          { series: "Tone", yearTarget: 0, monthly: [1010, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "蔡曜仲", group: "Div2", customers: [
      { name: "嘉義陽明", channel: "地區醫院", system: "", city: "嘉義市西區", products: [
          { series: "Ezole", yearTarget: 926478, monthly: [46743, 85695, 46743, 43200, 39126, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 110400, monthly: [12000, 8000, 8000, 8000, 8000, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 107000, monthly: [8533, 8533, 8533, 8533, 8533, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 100607, monthly: [32493, 0, 0, 30747, -1747, null, null, null, null, null, null, null] }
      ] },
      { name: "信安醫院", channel: "地區醫院", system: "", city: "雲林縣斗六市", products: [
          { series: "Ezole", yearTarget: 156880, monthly: [16952, 12714, 12714, 0, 12429, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 139212, monthly: [12381, 20952, 14762, 13571, 16667, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 127999, monthly: [14667, 5600, 5600, 14667, 5600, null, null, null, null, null, null, null] },
          { series: "Cospirit", yearTarget: 67982, monthly: [48571, 32381, 32381, 0, 24000, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 60150, monthly: [12029, 8020, 8020, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [16238, 0, 5600, 0, 16238, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 16514, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "嘉義基督教", channel: "區域醫院", system: "嘉基", city: "嘉義市東區", products: [
          { series: "Epine", yearTarget: 367070, monthly: [21429, 28571, 28572, 42858, 21429, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 145080, monthly: [11200, 11200, 22400, 11200, 11200, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 113050, monthly: [13295, 13295, 0, 13295, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部嘉義", channel: "地區醫院", system: "部標", city: "嘉義市東區", products: [
          { series: "Lote", yearTarget: 476280, monthly: [87745, 28305, 35381, 56609, 28305, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [16108, 0, 11275, 9126, 3758, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 6393, 0, 0, 6393, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 0, 0, 5267, null, null, null, null, null, null, null] }
      ] },
      { name: "嘉義榮民", channel: "區域醫院", system: "中榮", city: "嘉義市西區", products: [
          { series: "Fute", yearTarget: 351230, monthly: [34110, 68220, 0, 34110, 34110, null, null, null, null, null, null, null] }
      ] },
      { name: "聖馬爾定", channel: "區域醫院", system: "", city: "嘉義市西區", products: [
          { series: "Lote", yearTarget: 221520, monthly: [35266, 0, 30228, 15114, 15114, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 4848, monthly: [3029, 4846, 6663, 6938, 3469, null, null, null, null, null, null, null] }
      ] },
      { name: "台大雲林", channel: "區域醫院", system: "台大", city: "雲林縣斗六市", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [54952, 34381, 34381, 54952, 34381, null, null, null, null, null, null, null] }
      ] },
      { name: "灣橋榮民", channel: "地區醫院", system: "中榮", city: "嘉義縣竹崎鄉", products: [
          { series: "Fute", yearTarget: 129580, monthly: [23873, 0, 20464, 10232, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "慶昇醫院", channel: "地區醫院", system: "", city: "嘉義市西區", products: [
          { series: "Epine", yearTarget: 63650, monthly: [16761, 0, 6704, 6704, 6704, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 39450, monthly: [7886, 2629, 2629, 2629, 5258, null, null, null, null, null, null, null] }
      ] },
      { name: "盧亞人醫院", channel: "地區醫院", system: "", city: "嘉義市西區", products: [
          { series: "Bisadyl Supp", yearTarget: 45612, monthly: [3619, 0, 0, 2171, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 24350, monthly: [7343, 1695, 3390, 0, 0, null, null, null, null, null, null, null] },
          { series: "Pane", yearTarget: 6954, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "中國北港", channel: "區域醫院", system: "中國", city: "雲林縣北港鎮", products: [
          { series: "Lote", yearTarget: 32670, monthly: [17829, 7924, 0, 9905, 5943, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 10500, monthly: [2095, 0, 0, 2095, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安心醫院", channel: "地區醫院", system: "", city: "嘉義市西區", products: [
          { series: "Bisadyl Supp", yearTarget: 18100, monthly: [18095, 0, 0, 10857, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "賴政哲", group: "Div2", customers: [
      { name: "高市凱旋", channel: "區域醫院", system: "高市(部標)", city: "高雄市苓雅區", products: [
          { series: "Lote", yearTarget: 605640, monthly: [32558, 33975, 28311, 22650, 28311, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [76706, 76706, 76706, 76706, 76706, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 193880, monthly: [15714, 10476, 20952, 10476, 10476, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [8335, 5378, 5650, 5914, 4841, null, null, null, null, null, null, null] }
      ] },
      { name: "部屏東", channel: "區域醫院", system: "部標", city: "屏東縣屏東市", products: [
          { series: "Lote", yearTarget: 670320, monthly: [70761, 0, 70762, 28305, 28305, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 83840, monthly: [0, 10476, 0, 10476, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [2693, 2693, 4040, 4040, 2693, null, null, null, null, null, null, null] }
      ] },
      { name: "高雄榮總", channel: "醫學中心", system: "高榮", city: "高雄市左營區", products: [
          { series: "Tonex", yearTarget: 189775, monthly: [18425, 18425, 18118, 36850, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 116081, monthly: [0, 11267, 11267, 0, 10140, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [20000, 16000, 19999, 16000, 16000, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 44330, monthly: [0, 4029, 4029, 0, 4029, null, null, null, null, null, null, null] }
      ] },
      { name: "阮綜合醫院", channel: "區域醫院", system: "", city: "高雄市苓雅區", products: [
          { series: "Slatone", yearTarget: 267900, monthly: [38858, 0, 58286, 0, 19429, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [24714, 0, 16477, 16477, 16477, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 48600, monthly: [4857, 0, 12952, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "屏東榮總龍泉分院", channel: "地區醫院", system: "高榮", city: "屏東縣內埔鄉", products: [
          { series: "Lote", yearTarget: 291893, monthly: [0, 0, 28167, 0, 32674, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 52715, monthly: [14724, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 52500, monthly: [0, 4210, 6314, 0, 4210, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 4404, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 2400, 0, 1600, null, null, null, null, null, null, null] }
      ] },
      { name: "聖功醫院", channel: "地區醫院", system: "", city: "高雄市苓雅區", products: [
          { series: "Epram", yearTarget: 153010, monthly: [17143, 8571, 28572, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 45870, monthly: [0, 4171, 0, 0, 2781, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 36120, monthly: [7714, 0, 3857, 2571, 3857, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [7010, 0, 0, 7010, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 12920, monthly: [3048, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Pane", yearTarget: 9263, monthly: [10714, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "屏東榮總", channel: "地區醫院", system: "高榮", city: "屏東縣屏東市", products: [
          { series: "Lote", yearTarget: 73255, monthly: [0, 0, 11267, 0, 5633, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 48300, monthly: [15786, 0, 11576, 15786, 11576, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [3201, 0, 2800, 1600, 2000, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 4030, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 551, 1927, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "高市民生", channel: "地區醫院", system: "高市(部標)", city: "高雄市苓雅區", products: [
          { series: "Lote", yearTarget: 54390, monthly: [0, 14156, 14152, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 34986, monthly: [7144, 0, 10716, 0, 3186, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [10772, 0, 5386, 8080, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 15720, monthly: [5238, 0, 0, 0, 5238, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3196, 0, 3196, 0, 3196, null, null, null, null, null, null, null] }
      ] },
      { name: "輔英醫院", channel: "區域醫院", system: "", city: "屏東縣東港鎮", products: [
          { series: "Slatone", yearTarget: 110250, monthly: [8380, 8380, 4190, 8380, 8380, null, null, null, null, null, null, null] }
      ] },
      { name: "部恆春", channel: "地區醫院", system: "部標", city: "屏東縣恆春鎮", products: [
          { series: "Lote", yearTarget: 82320, monthly: [7076, 28304, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [4794, 6393, 0, 6393, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 2693, 2693, 0, 2693, null, null, null, null, null, null, null] }
      ] },
      { name: "部台東", channel: "地區醫院", system: "部標", city: "台東縣台東市", products: [
          { series: "Epine", yearTarget: 28024, monthly: [14629, 14629, 7314, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 2693, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "台東聖母", channel: "地區醫院", system: "", city: "台東縣台東市", products: [
          { series: "Lote", yearTarget: 13621, monthly: [12705, 12705, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3166, monthly: [4286, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 2502, monthly: [0, 3048, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2086, monthly: [1060, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部旗山", channel: "地區醫院", system: "部標", city: "高雄市旗山區", products: [
          { series: "Ezole", yearTarget: 7410, monthly: [0, 15378, 31324, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "台東基督教", channel: "地區醫院", system: "", city: "台東縣台東市", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [8228, 0, 6857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 9856, monthly: [8572, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 5650, monthly: [0, 2476, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [5143, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "屏東榮總榮譽之家", channel: "診所(含門前藥局)", system: "高榮", city: "屏東縣內埔鄉", products: [
          { series: "Lote", yearTarget: 32683, monthly: [2830, 2830, 2253, 2253, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [800, 0, 0, 400, 400, null, null, null, null, null, null, null] }
      ] },
      { name: "恆春基督教", channel: "地區醫院", system: "", city: "屏東縣恆春鎮", products: [
          { series: "Pane", yearTarget: 13755, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 12540, monthly: [0, 0, 0, 0, 2286, null, null, null, null, null, null, null] }
      ] },
      { name: "佑青醫院", channel: "地區醫院", system: "", city: "屏東縣內埔鄉", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [1619, 3238, 1619, 1619, 1619, null, null, null, null, null, null, null] }
      ] },
      { name: "博田國際醫院", channel: "地區醫院", system: "", city: "高雄市左營區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [3760, 0, 1791, 0, 2327, null, null, null, null, null, null, null] }
      ] },
      { name: "台東榮民", channel: "地區醫院", system: "北榮", city: "台東縣台東市", products: [
          { series: "Lote", yearTarget: 4032, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 2286, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 1208, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "祐生醫院", channel: "地區醫院", system: "", city: "高雄市三民區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 6705, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "民眾醫院", channel: "地區醫院", system: "", city: "屏東縣屏東市", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 1508, 2514, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "健新醫院", channel: "地區醫院", system: "", city: "高雄市前金區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [960, 960, 960, 0, 480, null, null, null, null, null, null, null] }
      ] },
      { name: "小港安泰", channel: "地區醫院", system: "", city: "高雄市小港區", products: [
          { series: "Fute", yearTarget: 1770, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "曾芷澄", group: "Div2", customers: [
      { name: "慈濟花蓮", channel: "醫學中心", system: "慈濟", city: "花蓮縣花蓮市", products: [
          { series: "Lote", yearTarget: 711942, monthly: [0, 35429, 35429, 35429, 62000, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 181302, monthly: [12714, 4714, 17428, 17428, 17428, null, null, null, null, null, null, null] }
      ] },
      { name: "振興醫院", channel: "區域醫院", system: "", city: "台北市北投區", products: [
          { series: "Bisadyl Supp", yearTarget: 238400, monthly: [24000, 12000, 12000, 12000, 12000, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 220000, monthly: [65734, 18781, 56343, 37562, 37562, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [35315, 35315, 35315, 35315, 35315, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [16190, 8095, 8095, 8095, 8095, null, null, null, null, null, null, null] }
      ] },
      { name: "台大醫院", channel: "醫學中心", system: "台大", city: "台北市中正區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [115427, 115427, 135998, 169141, 135998, null, null, null, null, null, null, null] }
      ] },
      { name: "玉里榮民", channel: "地區醫院", system: "北榮", city: "花蓮縣玉里鎮", products: [
          { series: "Fute", yearTarget: 112529, monthly: [34095, 0, 34095, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 85905, monthly: [42743, 85486, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 71284, monthly: [55689, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "悅思身心診所", channel: "診所(含門前藥局)", system: "", city: "花蓮縣吉安鄉", products: [
          { series: "Ezole", yearTarget: 144018, monthly: [11428, 0, 11428, 10857, 10857, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 73530, monthly: [0, 0, 12857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 64000, monthly: [8000, 0, 4000, 16191, 16190, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 39864, monthly: [0, 0, 0, 0, 6667, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 38688, monthly: [0, 0, 12476, 12476, 12476, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 1048, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市關渡", channel: "地區醫院", system: "特殊", city: "台北市北投區", products: [
          { series: "Epram", yearTarget: 129870, monthly: [11715, 5857, 17572, 11715, 5857, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [12085, 0, 12085, 8057, 8056, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 39675, monthly: [6459, 0, 3229, 3229, 3229, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [4628, 4628, 4628, 4628, 4628, null, null, null, null, null, null, null] }
      ] },
      { name: "台北榮總", channel: "醫學中心", system: "北榮本院", city: "台北市北投區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [40286, 20143, 60429, 20143, 40286, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 11571, 23142, 11571, 34714, null, null, null, null, null, null, null] }
      ] },
      { name: "部八里療", channel: "區域醫院", system: "部標", city: "新北市八里區", products: [
          { series: "Lote", yearTarget: 64680, monthly: [7076, 0, 4245, 7076, 7076, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 57640, monthly: [0, 10476, 0, 10476, 10476, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 1347, 1347, 1347, 2693, null, null, null, null, null, null, null] }
      ] },
      { name: "宏慈療養院", channel: "地區醫院", system: "宏慈", city: "新北市新店區", products: [
          { series: "Fute", yearTarget: 71775, monthly: [0, 13048, 13048, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 24435, monthly: [0, 0, 0, 3619, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 13716, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 8666, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部花蓮", channel: "地區醫院", system: "部標", city: "花蓮縣花蓮市", products: [
          { series: "Epram", yearTarget: 98325, monthly: [8266, 8267, 0, 12401, 5167, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 3196, 3196, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 1347, 0, 1347, 1347, null, null, null, null, null, null, null] }
      ] },
      { name: "軍花蓮805", channel: "區域醫院", system: "軍標", city: "花蓮縣新城鄉", products: [
          { series: "Fute", yearTarget: 63660, monthly: [5305, 5305, 5305, 5305, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [6438, 9657, 9657, 8047, 8047, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1665, 1666, 1667, 1667, 1666, null, null, null, null, null, null, null] }
      ] },
      { name: "門諾醫院", channel: "區域醫院", system: "門諾", city: "花蓮縣花蓮市", products: [
          { series: "Tonex", yearTarget: 53265, monthly: [7349, 7349, 0, 7349, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 26200, monthly: [3931, 9173, 6552, 6552, 6552, null, null, null, null, null, null, null] }
      ] },
      { name: "部玉里", channel: "地區醫院", system: "部標", city: "花蓮縣玉里鎮", products: [
          { series: "Bisadyl Supp", yearTarget: 32190, monthly: [20047, 16038, 16038, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 13501, monthly: [7076, 4245, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [2148, 1073, 1610, 0, 0, null, null, null, null, null, null, null] },
          { series: "Pane", yearTarget: 3069, monthly: [0, 2590, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1643, monthly: [0, 1667, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "宏濟醫院", channel: "地區醫院", system: "宏慈", city: "新北市新店區", products: [
          { series: "Fute", yearTarget: 44820, monthly: [0, 13048, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 29865, monthly: [0, 0, 0, 7238, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 13866, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 4572, monthly: [0, 0, 3810, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "耕莘新店", channel: "區域醫院", system: "耕莘", city: "新北市新店區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [13334, 13334, 13334, 26668, 13334, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 18058, 0, 12038, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "慈濟玉里", channel: "地區醫院", system: "慈濟", city: "花蓮縣玉里鎮", products: [
          { series: "Lote", yearTarget: 65527, monthly: [5314, 7086, 5314, 7086, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 7065, monthly: [0, 0, 1571, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "耕莘安康", channel: "區域醫院", system: "耕莘", city: "新北市新店區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [10666, 0, 10666, 5333, 10666, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3009, 3009, 3009, 3009, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "耕莘永和", channel: "地區醫院", system: "耕莘", city: "新北市永和區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [10667, 0, 10667, 8000, 10667, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 1461, 4470, 0, 3009, null, null, null, null, null, null, null] }
      ] },
      { name: "和信治癌", channel: "區域醫院", system: "", city: "台北市北投區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [4685, 4685, 0, 4686, 4685, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 10968, monthly: [2743, 0, 0, 0, 2743, null, null, null, null, null, null, null] }
      ] },
      { name: "維康-信光藥局(新光醫院)", channel: "健保藥局", system: "1B", city: "台北市士林區", products: [
          { series: "Lote", yearTarget: 19776, monthly: [3402, 3402, 3509, 3509, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "遠東聯合診所(亞東)", channel: "診所(含門前藥局)", system: "亞東", city: "台北市中正區", products: [
          { series: "Tone", yearTarget: 0, monthly: [2476, 4952, 2476, 2476, 4952, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [2400, 1800, 3300, 2700, 1800, null, null, null, null, null, null, null] }
      ] },
      { name: "北市昆明", channel: "區域醫院", system: "北市", city: "台北市萬華區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [3391, 5086, 8477, 0, 5086, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1600, 0, 0, 1600, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市林森(慢性病)", channel: "區域醫院", system: "北市", city: "台北市中山區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [3391, 2712, 5086, 0, 5086, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1600, 0, 1600, 1600, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北市陽明", channel: "區域醫院", system: "北市", city: "台北市士林區", products: [
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 0, 5476, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 3200, 1600, null, null, null, null, null, null, null] }
      ] },
      { name: "鳳林榮民", channel: "地區醫院", system: "北榮", city: "花蓮縣鳳林鎮", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [1610, 0, 806, 806, 805, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [2777, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 2220, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "維康信光藥局(新光醫院)", channel: "健保藥局", system: "1B", city: "台北市士林區", products: [
          { series: "Lote", yearTarget: 6592, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "金英寶藥局(同心診所)", channel: "診所(含門前藥局)", system: "", city: "花蓮縣花蓮市", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [1905, 0, 2476, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "方振名", group: "Div2", customers: [
      { name: "部嘉南療", channel: "區域醫院", system: "部標", city: "台南市仁德區", products: [
          { series: "Fute", yearTarget: 738440, monthly: [52381, 69524, 69524, 0, 69524, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 161700, monthly: [0, 28305, 0, 0, 28305, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 39900, monthly: [22781, 45562, 28476, 26333, 26333, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 7516, 5386, 0, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 7350, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "部台南", channel: "區域醫院", system: "部標", city: "台南市中西區", products: [
          { series: "Fute", yearTarget: 256760, monthly: [15715, 20952, 15715, 26190, 10476, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [8080, 5386, 9426, 4040, 6733, null, null, null, null, null, null, null] }
      ] },
      { name: "麻豆新樓", channel: "地區醫院", system: "新樓", city: "台南市麻豆區", products: [
          { series: "Exprexa", yearTarget: 149488, monthly: [24000, 0, 9600, 7246, 12720, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 72500, monthly: [15028, 0, 7514, 0, 7514, null, null, null, null, null, null, null] }
      ] },
      { name: "部新營", channel: "地區醫院", system: "部標", city: "台南市新營區", products: [
          { series: "Bisadyl Supp", yearTarget: 98646, monthly: [8019, 12028, 6015, 8019, 4812, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 45570, monthly: [0, 7076, 4245, 0, 4245, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 31440, monthly: [0, 5238, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3196, 9588, 0, 6393, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [5386, 2693, 0, 5386, 2693, null, null, null, null, null, null, null] }
      ] },
      { name: "奇美永康", channel: "醫學中心", system: "奇美", city: "台南市永康區", products: [
          { series: "Epram", yearTarget: 152820, monthly: [10858, 10858, 16287, 10858, 10858, null, null, null, null, null, null, null] }
      ] },
      { name: "市立安南", channel: "區域醫院", system: "中國", city: "台南市安南區", products: [
          { series: "Lote", yearTarget: 148500, monthly: [9905, 9905, 0, 9904, 4952, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [7048, 0, 0, 3524, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "台南新樓", channel: "區域醫院", system: "新樓", city: "台南市東區", products: [
          { series: "Epram", yearTarget: 90000, monthly: [12524, 0, 10020, 5010, 10020, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 28778, monthly: [4800, 0, 0, 963, 1413, null, null, null, null, null, null, null] }
      ] },
      { name: "成大醫院", channel: "醫學中心", system: "成大", city: "台南市北區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [34713, 34713, 19285, 38570, 30856, null, null, null, null, null, null, null] }
      ] },
      { name: "部朴子", channel: "地區醫院", system: "部標", city: "嘉義縣朴子市", products: [
          { series: "Epine", yearTarget: 53680, monthly: [4877, 6096, 4877, 3657, 4877, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 36680, monthly: [5238, 5238, 0, 0, 5238, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1074, 537, 537, 0, 2421, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2280, monthly: [1709, 2278, 0, 1054, 2634, null, null, null, null, null, null, null] }
      ] },
      { name: "台南榮民", channel: "地區醫院", system: "高榮", city: "台南市永康區", products: [
          { series: "Tonex", yearTarget: 64487, monthly: [7370, 11055, 7370, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [12000, 8000, 0, 8001, 16000, null, null, null, null, null, null, null] }
      ] },
      { name: "奇美柳營", channel: "區域醫院", system: "奇美", city: "台南市柳營區", products: [
          { series: "Epram", yearTarget: 32400, monthly: [4886, 0, 3258, 3258, 3258, null, null, null, null, null, null, null] }
      ] },
      { name: "奇美佳里", channel: "地區醫院", system: "奇美", city: "台南市佳里區", products: [
          { series: "Epram", yearTarget: 27270, monthly: [2172, 2171, 2172, 3258, 2172, null, null, null, null, null, null, null] }
      ] },
      { name: "成大斗六", channel: "地區醫院", system: "成大", city: "雲林縣斗六市", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [13500, 0, 6557, 7714, 7714, null, null, null, null, null, null, null] }
      ] },
      { name: "軍岡山814", channel: "地區醫院", system: "軍標", city: "高雄市岡山區", products: [
          { series: "Tone", yearTarget: 0, monthly: [9144, 3048, 4572, 4572, 4572, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "曾慧娥", group: "Div3", customers: [
      { name: "修慧診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣和美鎮", products: [
          { series: "Epine", yearTarget: 598349, monthly: [63332, 52381, 0, 35852, 42100, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 298480, monthly: [45868, 22934, 0, 17160, 22879, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 291938, monthly: [0, 32931, 0, 15840, 42103, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 258858, monthly: [26286, 39428, 0, 19714, 19714, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [22286, 22286, 7429, 5200, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [9191, 9191, 9191, 20219, 9191, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 51495, monthly: [15695, 0, 0, 1772, 12276, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 33696, monthly: [0, 0, 9714, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 3286, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "晨光藥局(心森林診所)", channel: "診所(含門前藥局)", system: "", city: "台中市北屯區", products: [
          { series: "Ritalin LA", yearTarget: 1049959, monthly: [131918, 60982, 89126, 38114, 73880, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 160665, monthly: [14320, 16171, 10000, 6171, 10000, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 52578, monthly: [15238, 7619, 0, 3810, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 28256, monthly: [2173, 0, 3621, 3621, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 26455, monthly: [0, 0, 0, 5714, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7428, 1486, 5200, 5943, 3715, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 3095, 3095, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "維新醫院", channel: "地區醫院", system: "", city: "台中市北區", products: [
          { series: "Epine", yearTarget: 397485, monthly: [27476, 34429, 43190, 46191, 53857, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 242604, monthly: [20933, 0, 20453, 8027, 25600, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 144996, monthly: [8990, 16857, 11238, 11048, 19886, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 129045, monthly: [0, 12286, 9829, 17200, 13514, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 111992, monthly: [0, 9080, 9080, 8960, 17920, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 96390, monthly: [21429, 0, 0, 17857, 3571, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 65051, monthly: [0, 0, 7120, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [16857, 0, 0, 16857, 20228, null, null, null, null, null, null, null] }
      ] },
      { name: "進化台安", channel: "地區醫院", system: "", city: "台中市東區", products: [
          { series: "Ritalin LA", yearTarget: 731826, monthly: [76446, 84484, 47044, 98856, 108051, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 277976, monthly: [36190, 27143, 24762, 21400, 24734, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14286, 21429, 21429, 17856, 8572, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 12954, monthly: [0, 0, 0, 0, 3805, null, null, null, null, null, null, null] }
      ] },
      { name: "和美身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣和美鎮", products: [
          { series: "Ritalin LA", yearTarget: 324591, monthly: [58400, 33143, 50514, 31543, 121543, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 291200, monthly: [28571, 0, 30286, 11429, 11429, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 136088, monthly: [23840, 0, 8160, 8629, 8160, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 108192, monthly: [9857, 0, 17476, 0, 9857, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 35750, monthly: [2857, 0, 0, 2857, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5943, 0, 2229, 1114, 1486, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 6840, monthly: [1524, 0, 0, 1524, 1524, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [5200, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "好晴天身心診所", channel: "診所(含門前藥局)", system: "好晴天", city: "台中市潭子區", products: [
          { series: "Ritalin LA", yearTarget: 606492, monthly: [123600, 0, 41477, 54260, 78683, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [28572, 0, 7143, 7143, 1429, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 28514, monthly: [0, 0, 0, 0, 8071, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [5257, 0, 3505, 2804, 2804, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 12155, monthly: [0, 2858, 0, 0, 2858, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 8140, monthly: [0, 0, 1476, 0, 0, null, null, null, null, null, null, null] },
          { series: "Suculin", yearTarget: 2266, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "嘉樂藥局(博智診所)", channel: "診所(含門前藥局)", system: "嘉樂", city: "台中市北區", products: [
          { series: "Epine", yearTarget: 365560, monthly: [38476, 22095, 38476, 38476, 16381, null, null, null, null, null, null, null] },
          { series: "Suculin", yearTarget: 130977, monthly: [8320, 8320, 8320, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 78057, monthly: [4715, 0, 4715, 12184, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 25715, 12857, 12858, 12857, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 33718, monthly: [0, 0, 0, 8788, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 18590, monthly: [0, 0, 5714, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 10799, monthly: [0, 0, 5611, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1600, monthly: [0, 0, 0, 0, 2400, null, null, null, null, null, null, null] }
      ] },
      { name: "佳樂身心診所", channel: "診所(含門前藥局)", system: "", city: "台中市北屯區", products: [
          { series: "Ritalin LA", yearTarget: 526568, monthly: [47715, 40000, 40000, 61714, 47715, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 57920, monthly: [7524, 0, 8858, 0, 7514, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7429, 7429, 7429, 7429, 1486, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 10010, monthly: [2858, 0, 0, 1429, 2858, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 3714, monthly: [0, 1238, 0, 2476, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "佳境藥局(漸漸身心診所)", channel: "診所(含門前藥局)", system: "", city: "台中市東區", products: [
          { series: "Ritalin LA", yearTarget: 335773, monthly: [42759, 24647, 26490, 34372, 27291, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 86112, monthly: [6629, 9942, 5524, 9391, 8838, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 81313, monthly: [7999, 9333, 6666, 6666, 8000, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 54020, monthly: [2953, 4428, 1477, 4428, 1476, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14277, 7514, 15029, 9768, 1503, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 27170, monthly: [2858, 0, 2857, 2857, 2857, null, null, null, null, null, null, null] }
      ] },
      { name: "蕭芸嶙身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "台中市太平區", products: [
          { series: "Epine", yearTarget: 227689, monthly: [43428, 0, 21333, 0, 37590, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 93764, monthly: [14765, 0, 7383, 0, 25840, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 81510, monthly: [17143, 0, 11430, 0, 13714, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 39749, monthly: [9466, 0, 3786, 0, 7467, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 37752, monthly: [4286, 0, 1714, 0, 4971, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [15400, 0, 9100, 3500, 4200, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 7868, monthly: [1125, 0, 0, 0, 2248, null, null, null, null, null, null, null] }
      ] },
      { name: "黃淑琦診所", channel: "診所(含門前藥局)", system: "", city: "台中市北區", products: [
          { series: "Ritalin LA", yearTarget: 456960, monthly: [85330, 0, 6165, 74428, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 35286, 0, 9175, null, null, null, null, null, null, null] }
      ] },
      { name: "富丘藥局(佑芯身心診所)", channel: "診所(含門前藥局)", system: "", city: "台中市北屯區", products: [
          { series: "Fute", yearTarget: 236080, monthly: [7619, 12191, 10667, 14143, 25667, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 92098, monthly: [18381, 5334, 9524, 17524, 19714, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 32890, monthly: [0, 0, 0, 0, 8572, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 30452, monthly: [0, 1324, 3971, 6619, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 28492, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [4507, 4506, 1040, 4506, 9013, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 16635, monthly: [2560, 0, 1280, 2560, 2560, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 12600, monthly: [2095, 0, 0, 4190, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 2229, 743, 743, 743, null, null, null, null, null, null, null] }
      ] },
      { name: "小柚子藥局(張容毓親子診所)", channel: "診所(含門前藥局)", system: "", city: "彰化縣鹿港鎮", products: [
          { series: "Ritalin LA", yearTarget: 395032, monthly: [38844, 35540, 40730, 25959, 51286, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [5133, 4212, 3026, 7109, 1583, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2972, 2972, 2972, 2229, 3715, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 14392, monthly: [0, 1542, 0, 1542, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 1524, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "好事彰化身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣彰化市", products: [
          { series: "Ritalin LA", yearTarget: 133531, monthly: [17361, 4743, 12629, 49714, 25258, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 106167, monthly: [17476, 0, 9857, 0, 9857, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 67592, monthly: [6381, 0, 9143, 2286, 23047, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 54600, monthly: [2720, 0, 5440, 2157, 15430, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1114, 2229, 4086, 2229, 1486, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [5200, 0, 0, 1733, 3467, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 9295, monthly: [4286, 0, 0, 0, 2857, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 2280, monthly: [3048, 0, 1524, 1524, 3048, null, null, null, null, null, null, null] }
      ] },
      { name: "道周醫院", channel: "地區醫院", system: "", city: "彰化縣和美鎮", products: [
          { series: "Exelon Patch", yearTarget: 158964, monthly: [18343, 18342, 18343, 17726, 17725, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 149494, monthly: [45280, 0, 0, 30186, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 23411, monthly: [0, 5714, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 6571, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "趙夢麒診所", channel: "診所(含門前藥局)", system: "", city: "雲林縣虎尾鎮", products: [
          { series: "Epine", yearTarget: 235250, monthly: [49048, 0, 0, 47619, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 72755, monthly: [25067, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 53970, monthly: [15428, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "存寬診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣員林市", products: [
          { series: "Ezole", yearTarget: 135999, monthly: [19714, 0, 0, 19714, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 99333, monthly: [0, 20266, 0, 0, 20266, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 72250, monthly: [10172, 0, 6781, 13562, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 41600, monthly: [9524, 4572, 8000, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 3571, 0, 4286, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 4286, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3325, monthly: [0, 0, 1333, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "敦仁醫院", channel: "地區醫院", system: "", city: "彰化縣員林市", products: [
          { series: "Epine", yearTarget: 201850, monthly: [20495, 20495, 0, 20495, 12905, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 112980, monthly: [0, 0, 21524, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 20570, monthly: [0, 0, 4114, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 3353, 0, 3353, 1676, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2971, 0, 1114, 1114, 2600, null, null, null, null, null, null, null] }
      ] },
      { name: "鄭曜忠身心診所", channel: "診所(含門前藥局)", system: "", city: "台中市大里區", products: [
          { series: "Lote", yearTarget: 120060, monthly: [19571, 0, 26095, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 63874, monthly: [0, 19929, 0, 4757, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 43840, monthly: [0, 4952, 6000, 0, 4952, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 33345, monthly: [8571, 0, 0, 8571, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 20347, monthly: [0, 0, 4286, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 14857, 0, 3714, 1486, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 14941, monthly: [7467, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 8743, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 7599, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "晨心身心診所", channel: "診所(含門前藥局)", system: "", city: "台中市太平區", products: [
          { series: "Ritalin LA", yearTarget: 179506, monthly: [34229, 0, 37372, 0, 15600, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 57405, monthly: [6429, 10286, 0, 0, 6429, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [21857, 0, 7286, 3643, 1457, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 23310, monthly: [0, 20000, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 9002, monthly: [2571, 0, 2571, 0, 2571, null, null, null, null, null, null, null] }
      ] },
      { name: "晴明身心診所", channel: "診所(含門前藥局)", system: "", city: "雲林縣斗六市", products: [
          { series: "Ritalin LA", yearTarget: 223845, monthly: [0, 0, 143623, 36128, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 33150, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北屯好晴天藥局(北屯好晴天診所)", channel: "診所(含門前藥局)", system: "好晴天", city: "台中市北屯區", products: [
          { series: "Ritalin LA", yearTarget: 146114, monthly: [61799, 4608, 13825, 69551, 18434, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14286, 14286, 7143, 3571, 4286, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 17758, monthly: [0, 0, 0, 0, 9143, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3505, 1752, 1753, 5257, 3504, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 5950, monthly: [2381, 0, 0, 0, 2381, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1480, monthly: [0, 0, 0, 1476, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1430, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "佳佑診所", channel: "診所(含門前藥局)", system: "展新", city: "台中市北屯區", products: [
          { series: "Ritalin LA", yearTarget: 175339, monthly: [20735, 20735, 31680, 15840, 20734, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 52215, monthly: [5314, 5314, 3543, 5314, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [15200, 15200, 7600, 2280, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "禾悅身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "台中市大里區", products: [
          { series: "Ritalin LA", yearTarget: 241915, monthly: [13017, 39286, 0, 13017, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7429, 7429, 3714, 3714, 743, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 22903, monthly: [0, 0, 2476, 0, 2476, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 12860, monthly: [0, 1286, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 7997, monthly: [0, 1333, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "李光耀診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣田尾鄉", products: [
          { series: "Epine", yearTarget: 167115, monthly: [17143, 12857, 0, 8562, 28552, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 84633, monthly: [26666, 0, 6667, 0, 4801, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 8450, monthly: [2590, 0, 0, 1295, 1295, null, null, null, null, null, null, null] }
      ] },
      { name: "溫建文診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣鹿港鎮", products: [
          { series: "Exprexa", yearTarget: 273193, monthly: [0, 46667, 23333, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 19734, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "中大振福藥局(中大身心診所)", channel: "診所(含門前藥局)", system: "", city: "台中市太平區", products: [
          { series: "Ritalin LA", yearTarget: 137685, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 58992, monthly: [0, 24572, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 29162, monthly: [10027, 0, 0, 10646, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 27612, monthly: [9809, 0, 0, 0, 13324, null, null, null, null, null, null, null] },
          { series: "Bisadyl", yearTarget: 0, monthly: [0, 0, 13714, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 7939, monthly: [0, 1371, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 7865, monthly: [0, 0, 0, 0, 4286, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1400, 1400, null, null, null, null, null, null, null] }
      ] },
      { name: "日之林藥局(春森診所)", channel: "診所(含門前藥局)", system: "", city: "台中市東區", products: [
          { series: "Exelon", yearTarget: 117024, monthly: [38667, 6667, 9333, 0, 18668, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 99969, monthly: [21214, 8600, 0, 11246, -350, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [6248, 3124, 0, 6247, 2499, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 6171, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1430, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "享開心身心診所", channel: "診所(含門前藥局)", system: "開心房", city: "台中市豐原區", products: [
          { series: "Epine", yearTarget: 66420, monthly: [6476, 6476, 0, 6476, 15390, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 64795, monthly: [11418, 0, 4782, 5486, 15944, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 41453, monthly: [0, 7620, 0, 0, 9334, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 22040, monthly: [0, 3048, 0, 3047, 3047, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 19710, monthly: [2629, 2629, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5200, 0, 2229, 743, 743, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 1781, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂欣藥局(文心樂丞診所)", channel: "診所(含門前藥局)", system: "佳璋", city: "台中市北屯區", products: [
          { series: "Lote", yearTarget: 88608, monthly: [9981, 0, 7486, 13723, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 36100, monthly: [3809, 2857, 5714, 3809, 3810, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [6210, 3105, 12420, 6210, 6210, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 33147, monthly: [4191, 5714, 0, 1810, 2171, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 24682, monthly: [4628, 0, 3857, 0, 4628, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 3571, 5714, 2143, 714, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1430, monthly: [0, 0, 0, 1429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "伍杉藥局(晹信診所)", channel: "診所(含門前藥局)", system: "", city: "彰化縣彰化市", products: [
          { series: "Epine", yearTarget: 70300, monthly: [38096, 0, 0, 3810, 19048, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 60442, monthly: [12857, 0, 12857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 34155, monthly: [9905, 0, 0, 0, 9905, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3285, 0, 3285, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 520, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "瑞生藥局(瑞健診所)", channel: "診所(含門前藥局)", system: "", city: "台中市北屯區", products: [
          { series: "Exelon", yearTarget: 124084, monthly: [42160, 0, 28106, 10540, 14053, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 28171, monthly: [0, 3200, 3086, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 5916, monthly: [0, 0, 1691, 0, 845, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 0, 4594, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [737, 0, 1474, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "品安藥局(陸建民診所)", channel: "診所(含門前藥局)", system: "", city: "台中市豐原區", products: [
          { series: "Exelon Patch", yearTarget: 78000, monthly: [7500, 0, 0, 14771, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [25905, 0, 0, 38858, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 30489, monthly: [3453, 0, 3454, 6907, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 30400, monthly: [5714, 0, 5714, 5714, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "卓醫院", channel: "地區醫院", system: "", city: "彰化縣北斗鎮", products: [
          { series: "Exelon Patch", yearTarget: 162041, monthly: [0, 0, 30286, 0, 29600, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 4929, 9856, 4929, 4929, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 3689, monthly: [3690, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 2286, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "德美藥局(和興聯合診所)", channel: "診所(含門前藥局)", system: "", city: "彰化縣和美鎮", products: [
          { series: "Exelon", yearTarget: 76430, monthly: [5728, 5011, 0, 10022, 14320, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 48571, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 21360, monthly: [0, 0, 2482, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 4572, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2668, monthly: [0, 0, 666, 0, 666, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 2400, monthly: [0, 2399, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 2380, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "王家駿身心診所", channel: "診所(含門前藥局)", system: "", city: "台中市北屯區", products: [
          { series: "Epine", yearTarget: 89394, monthly: [5190, 9190, 5190, 5190, 5190, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 20815, monthly: [0, 0, 1810, 1810, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7600, 0, 7600, 760, 2280, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 4725, monthly: [0, 1048, 0, 1048, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "雙十台安", channel: "地區醫院", system: "", city: "台中市北區", products: [
          { series: "Ritalin LA", yearTarget: 84096, monthly: [0, 0, 7125, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 33350, monthly: [4071, 0, 2714, 2714, 1357, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [2419, 2419, 2419, 2419, 2419, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 7955, monthly: [0, 0, 3977, 0, 0, null, null, null, null, null, null, null] },
          { series: "Pane", yearTarget: 2303, monthly: [1371, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [599, 0, 0, 0, 374, null, null, null, null, null, null, null] }
      ] },
      { name: "佑仁診所", channel: "診所(含門前藥局)", system: "", city: "雲林縣二崙鄉", products: [
          { series: "Exelon Patch", yearTarget: 58061, monthly: [0, 0, 3143, 12068, 7400, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 24762, 24762, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 18816, monthly: [1466, 4400, 2933, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 1585, 1057, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "王志中診所", channel: "診所(含門前藥局)", system: "", city: "台中市大里區", products: [
          { series: "Ezole", yearTarget: 65688, monthly: [0, 9524, 0, 9524, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 43500, monthly: [0, 7524, 0, 0, 7524, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 14564, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "平安藥局(吳潮聰診所)", channel: "診所(含門前藥局)", system: "", city: "彰化縣彰化市", products: [
          { series: "Fute", yearTarget: 68870, monthly: [0, 3048, 9905, 3047, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 21022, monthly: [0, 3657, 0, 3658, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 20900, monthly: [2857, 2857, 0, 2857, 2857, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2440, 0, 939, 939, 626, null, null, null, null, null, null, null] }
      ] },
      { name: "小鹿藥局(品心診所)", channel: "診所(含門前藥局)", system: "", city: "台中市東區", products: [
          { series: "Exelon", yearTarget: 87934, monthly: [12960, 0, 0, 8640, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 11400, monthly: [2286, 0, 0, 3428, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 7857, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 3200, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2191, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "世淋診所", channel: "診所(含門前藥局)", system: "", city: "台中市豐原區", products: [
          { series: "Ezole", yearTarget: 66660, monthly: [20190, 0, 20192, 1514, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 15178, monthly: [2371, 0, 2371, 2371, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [372, 0, 0, 743, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "福安醫院", channel: "地區醫院", system: "", city: "雲林縣斗南鎮", products: [
          { series: "Exelon", yearTarget: 60553, monthly: [14933, 14933, 3733, 7467, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3943, 1314, 2629, 1314, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 0, 8172, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 4114, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [714, 1071, 714, 714, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [485, 485, 485, 485, 485, null, null, null, null, null, null, null] }
      ] },
      { name: "明功堂診所", channel: "診所(含門前藥局)", system: "", city: "台中市東區", products: [
          { series: "Fute", yearTarget: 52550, monthly: [10381, 0, 0, 0, 6095, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 35245, monthly: [0, 0, 5572, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 10511, monthly: [0, 0, 4572, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3800, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "東仁藥局(廖寶全診所)", channel: "診所(含門前藥局)", system: "", city: "雲林縣虎尾鎮", products: [
          { series: "Exelon Patch", yearTarget: 77990, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 19594, monthly: [0, 915, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1971, 2629, 0, 0, 3286, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 3714, 0, 1486, 1783, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 6934, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "心仁藥局(明朗身心診所)", channel: "診所(含門前藥局)", system: "", city: "台中市大里區", products: [
          { series: "Ezole", yearTarget: 19305, monthly: [8572, 0, 0, 8572, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7429, 7429, 2229, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 16833, monthly: [0, 4646, 0, 3872, 3872, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 16480, monthly: [0, 0, 0, 0, 2567, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 9002, monthly: [3857, 0, 0, 0, 3857, null, null, null, null, null, null, null] }
      ] },
      { name: "何正岳診所", channel: "診所(含門前藥局)", system: "", city: "雲林縣斗六市", products: [
          { series: "Epine", yearTarget: 41920, monthly: [7858, 0, 7858, 0, 7857, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 17857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 17526, monthly: [3809, 0, 3809, 0, 7619, null, null, null, null, null, null, null] }
      ] },
      { name: "雲萱藥局(雲萱診所)", channel: "診所(含門前藥局)", system: "", city: "雲林縣斗六市", products: [
          { series: "Ezole", yearTarget: 65688, monthly: [0, 0, 14285, 0, 14286, null, null, null, null, null, null, null] }
      ] },
      { name: "智樂藥局(理解身心診所)", channel: "診所(含門前藥局)", system: "佳璋", city: "台中市東區", products: [
          { series: "Lote", yearTarget: 16224, monthly: [3743, 0, 0, 2495, 2495, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 13716, monthly: [2286, 0, 0, 0, 1810, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7500, 0, 2143, 714, 714, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 9025, monthly: [1904, 2857, 0, 2857, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [1863, 0, 2484, 931, 2484, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 6170, monthly: [3085, 0, 0, 0, 4628, null, null, null, null, null, null, null] }
      ] },
      { name: "悅康診所", channel: "診所(含門前藥局)", system: "", city: "台中市大里區", products: [
          { series: "Suculin", yearTarget: 39984, monthly: [4000, 0, 4000, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [6571, 0, 3285, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 0, monthly: [0, 4114, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 2380, monthly: [0, 0, 2381, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1710, monthly: [0, 0, 1715, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "建元醫院", channel: "地區醫院", system: "", city: "彰化縣田中鎮", products: [
          { series: "Fute", yearTarget: 30000, monthly: [10000, 0, 0, 0, 10000, null, null, null, null, null, null, null] }
      ] },
      { name: "賢德醫院", channel: "地區醫院", system: "賢德", city: "台中市太平區", products: [
          { series: "Epine", yearTarget: 19440, monthly: [7286, 0, 0, 2424, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [8017, 0, 0, 1146, 1833, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1520, 0, 0, 760, 1140, null, null, null, null, null, null, null] }
      ] },
      { name: "霧峰澄清", channel: "地區醫院", system: "", city: "台中市大里區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [6915, 6915, 0, 6915, 6915, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [5514, 0, 0, 5514, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新太平澄清", channel: "地區醫院", system: "", city: "台中市太平區", products: [
          { series: "Tonex", yearTarget: 23802, monthly: [0, 7324, 0, 0, 7324, null, null, null, null, null, null, null] }
      ] },
      { name: "福瑞藥局(杏安診所)", channel: "診所(含門前藥局)", system: "", city: "雲林縣虎尾鎮", products: [
          { series: "Exelon", yearTarget: 13366, monthly: [1532, 1532, 0, 3064, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 7715, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 1715, 0, 1714, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [2305, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 1584, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 293, null, null, null, null, null, null, null] }
      ] },
      { name: "康森藥局(卓立復健科診所)", channel: "診所(含門前藥局)", system: "", city: "台中市大里區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [8229, 0, 8229, 8229, 0, null, null, null, null, null, null, null] },
          { series: "Suculin", yearTarget: 8197, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "敦愛身心診所", channel: "診所(含門前藥局)", system: "", city: "台中市北屯區", products: [
          { series: "Ezole", yearTarget: 10948, monthly: [0, 1428, 952, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5005, monthly: [0, 1429, 0, 0, 1428, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 3240, monthly: [0, 0, 2905, 0, 1286, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1430, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 1280, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "雄安骨科診所", channel: "診所(含門前藥局)", system: "", city: "台中市大里區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 18447, 0, 11068, null, null, null, null, null, null, null] }
      ] },
      { name: "員林健康藥局(高杏診所)", channel: "診所(含門前藥局)", system: "", city: "彰化縣員林市", products: [
          { series: "Epine", yearTarget: 11520, monthly: [0, 3905, 2533, 0, 2533, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3800, monthly: [0, 0, 1905, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "宏科藥局(張右川診所)", channel: "診所(含門前藥局)", system: "", city: "台中市北區", products: [
          { series: "Tone", yearTarget: 18240, monthly: [0, 3428, 0, 3428, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "喜達康藥局(陳建達診所)", channel: "診所(含門前藥局)", system: "", city: "彰化縣北斗鎮", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [4543, 0, 1514, 1514, 757, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 7015, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3575, monthly: [0, 0, 2857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [1238, 0, 1238, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林翰宏骨科診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣和美鎮", products: [
          { series: "Suculin", yearTarget: 17369, monthly: [2053, 2053, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "聯安醫院", channel: "地區醫院", system: "", city: "台中市北屯區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [4857, 4857, 0, 4857, 4857, null, null, null, null, null, null, null] }
      ] },
      { name: "宏佑診所", channel: "診所(含門前藥局)", system: "", city: "雲林縣斗南鎮", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 6142, 0, 0, 12286, null, null, null, null, null, null, null] }
      ] },
      { name: "頤晴診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "彰化縣彰化市", products: [
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 7619, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3500, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3100, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "優生婦產科診所", channel: "診所(含門前藥局)", system: "", city: "台中市豐原區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 0, 13494, null, null, null, null, null, null, null] }
      ] },
      { name: "中貞藥局(元亨診所)", channel: "診所(含門前藥局)", system: "", city: "台中市北屯區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [4251, 0, 4251, 0, 4251, null, null, null, null, null, null, null] }
      ] },
      { name: "宗生診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣秀水鄉", products: [
          { series: "Exelon Patch", yearTarget: 9171, monthly: [0, 0, 0, 2965, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "劉騰光診所", channel: "診所(含門前藥局)", system: "", city: "台中市北區", products: [
          { series: "Ezole", yearTarget: 6188, monthly: [0, 0, 2857, 0, 2857, null, null, null, null, null, null, null] }
      ] },
      { name: "百禾藥局中清店(佑康診所)", channel: "診所(含門前藥局)", system: "", city: "台中市北屯區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 6523, 0, 0, 4894, null, null, null, null, null, null, null] }
      ] },
      { name: "青田藥局(張天長診所)", channel: "診所(含門前藥局)", system: "", city: "彰化縣員林市", products: [
          { series: "Tonex", yearTarget: 0, monthly: [4114, 0, 0, 0, 4114, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 2666, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "伸港忠孝", channel: "地區醫院", system: "伸港", city: "彰化縣伸港鄉", products: [
          { series: "Exelon", yearTarget: 0, monthly: [0, 0, 0, 16908, -8454, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 1562, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "合康藥局(合濟診所)", channel: "診所(含門前藥局)", system: "", city: "彰化縣北斗鎮", products: [
          { series: "Tone", yearTarget: 9880, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "人愛診所", channel: "診所(含門前藥局)", system: "伸港", city: "台中市潭子區", products: [
          { series: "Fute", yearTarget: 8145, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 1143, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "益嘉藥局(佳藝診所)", channel: "診所(含門前藥局)", system: "", city: "台中市豐原區", products: [
          { series: "Epine", yearTarget: 4290, monthly: [0, 0, 2857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 2100, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "郁順復健科診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣員林市", products: [
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 4114, 0, 4114, null, null, null, null, null, null, null] }
      ] },
      { name: "蕭永明骨科診所", channel: "診所(含門前藥局)", system: "", city: "台中市北屯區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [0, 3689, 0, 0, 3689, null, null, null, null, null, null, null] }
      ] },
      { name: "黃明弘診所", channel: "診所(含門前藥局)", system: "", city: "台中市豐原區", products: [
          { series: "Epine", yearTarget: 3420, monthly: [3429, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "芳鄰診所", channel: "診所(含門前藥局)", system: "", city: "台中市北屯區", products: [
          { series: "Tone", yearTarget: 3675, monthly: [0, 0, 0, 3143, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "杏豐診所", channel: "診所(含門前藥局)", system: "", city: "台中市豐原區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [3353, 0, 0, 0, 3353, null, null, null, null, null, null, null] }
      ] },
      { name: "陳彥學內科診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣員林市", products: [
          { series: "Tone", yearTarget: 6300, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "陽光小兒科診所", channel: "診所(含門前藥局)", system: "", city: "彰化縣彰化市", products: [
          { series: "Exelon", yearTarget: 5944, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "智悅藥局(智欣診所)", channel: "診所(含門前藥局)", system: "嘉樂", city: "台中市西區", products: [
          { series: "Ritalin LA", yearTarget: 5601, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "本堂澄清", channel: "地區醫院", system: "", city: "台中市霧峰區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [3505, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "蔡醫院", channel: "地區醫院", system: "", city: "雲林縣土庫鎮", products: [
          { series: "Epine", yearTarget: 3430, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "東亞藥師大藥局(黃啟文診所)", channel: "診所(含門前藥局)", system: "", city: "台中市太平區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 0, 2482, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 737, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "勝美醫院", channel: "地區醫院", system: "保健安", city: "台中市北區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [1971, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "周宇君", group: "Div3", customers: [
      { name: "千仟藥局(林正修診所)", channel: "診所(含門前藥局)", system: "", city: "新竹市東區", products: [
          { series: "Ritalin LA", yearTarget: 2238706, monthly: [313097, 0, 281828, 31269, 105531, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [63429, 0, 63429, 0, 8245, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 122642, monthly: [29619, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 60417, monthly: [10080, 0, 12880, 0, 9813, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 25585, monthly: [12038, 0, 11495, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 9940, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "黃正龍診所", channel: "診所(含門前藥局)", system: "", city: "桃園市平鎮區", products: [
          { series: "Ritalin LA", yearTarget: 1224563, monthly: [112543, 34586, 147130, 91443, 126029, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 590000, monthly: [30000, 30000, 50000, 50000, 60000, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 246563, monthly: [0, 8191, 20477, 20477, 40953, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [9857, 9857, 13142, 1971, 3285, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 0, monthly: [0, 0, 2020, 2020, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "普羅旺斯專科藥局(能清安欣診所)", channel: "診所(含門前藥局)", system: "", city: "新竹市東區", products: [
          { series: "Ritalin LA", yearTarget: 1090391, monthly: [104370, 111455, 91075, 89791, 114791, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 110713, monthly: [7115, 9742, 7515, 0, 3170, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14520, 18216, 16500, 9900, 660, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 47568, monthly: [5949, 7933, 1983, 0, 4957, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 23520, monthly: [2514, 3352, 1676, 0, 2514, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 20198, monthly: [1922, 0, 1923, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 10220, monthly: [1143, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 9975, monthly: [1143, 1143, 4381, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "全民好藥局(馬大元診所)", channel: "診所(含門前藥局)", system: "", city: "新竹市東區", products: [
          { series: "Ritalin LA", yearTarget: 877606, monthly: [0, 87395, 153814, 0, 14502, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 136620, monthly: [0, 0, 13809, 0, 13809, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 85100, monthly: [0, 0, 8429, 0, 22857, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 75210, monthly: [3271, 5452, 32714, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 72590, monthly: [0, 4762, 11905, 0, 11905, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 61542, monthly: [6571, 0, 15144, 8762, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [21429, 0, 21429, 0, 2857, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [16096, 0, 0, 0, 9657, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [8667, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 932, monthly: [1867, 0, 0, 2680, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "湖口仁慈", channel: "地區醫院", system: "耕莘", city: "新竹縣湖口鄉", products: [
          { series: "Exelon Patch", yearTarget: 448995, monthly: [29360, 52169, 30411, 44486, 57074, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 245737, monthly: [12597, 9448, 18896, 9448, 6299, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 103286, monthly: [12010, 12010, 12010, 0, 6005, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 68275, monthly: [8533, 4267, 4267, 0, 4267, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [5438, 5438, 21752, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [6667, 6667, 6667, 6667, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3009, 3009, 6019, 6019, 6019, null, null, null, null, null, null, null] },
          { series: "Ramesoon", yearTarget: 0, monthly: [0, 7200, 0, 7200, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5000, 0, 5000, 2143, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "親禾身心診所", channel: "診所(含門前藥局)", system: "平衡", city: "新竹市東區", products: [
          { series: "Ritalin LA", yearTarget: 479479, monthly: [68580, 28529, 77124, 45354, 82474, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 122754, monthly: [17572, 12572, 11786, 6286, 24667, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 67151, monthly: [7600, 0, 0, 0, 7600, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 57405, monthly: [7714, 0, 7714, 0, 5143, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7229, 0, 10843, 2169, 2892, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 17750, monthly: [4762, 0, 4761, 1714, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "尚語身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市中壢區", products: [
          { series: "Ritalin LA", yearTarget: 435786, monthly: [0, 0, 53053, 37140, 53053, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 214828, monthly: [0, 0, 0, 0, 62429, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [44571, 0, 14857, 7429, 1486, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 60088, monthly: [4725, 0, 15406, 0, 14658, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 22858, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 6115, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "芮成藥局(平衡身心診所)", channel: "診所(含門前藥局)", system: "平衡", city: "新竹市東區", products: [
          { series: "Ritalin LA", yearTarget: 316183, monthly: [83563, 0, 18150, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 155624, monthly: [57071, 0, 0, 0, 18857, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 81270, monthly: [19286, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 69685, monthly: [21533, 0, 21533, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 62515, monthly: [11238, 0, 11238, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14458, 0, 5783, 0, 5061, null, null, null, null, null, null, null] }
      ] },
      { name: "暖林藥局(暖陽身心診所)", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹北市", products: [
          { series: "Epine", yearTarget: 178320, monthly: [32532, 13867, 13867, 13676, 13676, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 164220, monthly: [11905, 11904, 11905, 11904, 11904, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 161466, monthly: [11597, 0, 11597, 11597, 20797, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 113505, monthly: [17296, 8648, 5404, 6486, 6486, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 14286, 3571, 3571, 3572, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 13485, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 6210, monthly: [0, 1381, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 1904, monthly: [0, 476, 0, 0, 953, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 1734, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新楊梅診所", channel: "診所(含門前藥局)", system: "", city: "桃園市楊梅區", products: [
          { series: "Ritalin LA", yearTarget: 378545, monthly: [25569, 29646, 19887, 48333, 88476, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 89880, monthly: [8665, 1762, 3819, 9257, 7542, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [11143, 22286, 11143, 5200, 2229, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 47304, monthly: [6134, 3504, 3504, 4190, 4190, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 32890, monthly: [2858, 1429, 2858, 2858, 2858, null, null, null, null, null, null, null] }
      ] },
      { name: "福德身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市中壢區", products: [
          { series: "Ritalin LA", yearTarget: 309291, monthly: [0, 0, 0, 0, 154286, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 192192, monthly: [0, 0, 0, 12371, 74229, null, null, null, null, null, null, null] },
          { series: "Cospirit", yearTarget: 0, monthly: [0, 0, 10667, 0, 10571, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 16783, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Suculin", yearTarget: 15406, monthly: [8800, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [13143, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 5200, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "聊癒之森身心診所", channel: "診所(含門前藥局)", system: "", city: "新竹市東區", products: [
          { series: "Ritalin LA", yearTarget: 305264, monthly: [121301, 0, 52793, 24194, 49173, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 61828, monthly: [8533, 0, 5333, 3086, 3086, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [28914, 0, 14457, 7229, 5061, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 36743, monthly: [3800, 0, 5067, 5067, 8866, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 19846, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 16610, monthly: [3019, 0, 3019, 1486, 7676, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [4037, 0, 1552, 3726, 1862, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3150, monthly: [1257, 0, 1257, 0, 1257, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [544, 0, 543, 542, 543, null, null, null, null, null, null, null] }
      ] },
      { name: "九禾藥局(陳炯旭診所)", channel: "診所(含門前藥局)", system: "", city: "桃園市中壢區", products: [
          { series: "Ritalin LA", yearTarget: 226508, monthly: [0, 21754, 0, 23525, 22557, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 157954, monthly: [19572, 22029, 0, 11742, 40829, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 70224, monthly: [7746, 11933, 0, 10494, 42560, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 7286, 2186, 2186, 1458, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [515, 0, 1029, 514, 515, null, null, null, null, null, null, null] }
      ] },
      { name: "安立身心診所", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 394670, monthly: [48133, 12532, 24513, 36328, 53766, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 64300, monthly: [2572, 2571, 3857, 5143, 5143, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 25705, monthly: [1600, 0, 5600, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [1369, 704, 1720, 1057, 976, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 0, monthly: [0, 0, 1143, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "周伯翰身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 285860, monthly: [49142, 4572, 32286, 13715, 38286, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 58465, monthly: [7071, 7071, 0, 11338, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 51246, monthly: [6571, 0, 6571, 0, 6571, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [7715, 19286, 0, 7714, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 14857, 7428, 5943, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1240, monthly: [0, 0, 1238, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "黃煌富身心診所", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 114707, monthly: [14708, 0, 14708, 14708, 14708, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 90440, monthly: [14286, 0, 14286, 14286, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 61568, monthly: [7186, 3238, 11133, 6476, 14095, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 47040, monthly: [5619, 3372, 5619, 5619, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 9660, monthly: [4143, 0, 5524, 5524, 5524, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 0, 1714, null, null, null, null, null, null, null] }
      ] },
      { name: "台齡身心診所", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹北市", products: [
          { series: "Ezole", yearTarget: 125060, monthly: [13524, 13524, 0, 13524, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 66960, monthly: [0, 6190, 3714, 3714, 6191, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 60072, monthly: [5760, 0, 5760, 0, 4937, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 34892, monthly: [0, 0, 6709, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [8674, 0, 0, 5783, 723, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [2446, 2446, 0, 2446, 2445, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 2850, monthly: [0, 0, 0, 0, 1143, null, null, null, null, null, null, null] }
      ] },
      { name: "中壢長榮", channel: "地區醫院", system: "", city: "桃園市中壢區", products: [
          { series: "Exelon", yearTarget: 234926, monthly: [13344, 0, 13344, 26688, 40032, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 44944, monthly: [3099, 0, 4650, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 6571, 0, 6571, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [380, 0, 380, 380, 380, null, null, null, null, null, null, null] }
      ] },
      { name: "培靈關西", channel: "地區醫院", system: "", city: "新竹縣關西鎮", products: [
          { series: "Exprexa", yearTarget: 109652, monthly: [0, 0, 7600, 11333, 0, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 90664, monthly: [0, 0, 13600, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 60169, monthly: [0, 0, 0, 5238, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 42400, monthly: [1600, 0, 9486, 0, 7714, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 6795, monthly: [0, 0, 1514, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "長慎醫院", channel: "地區醫院", system: "安慎", city: "桃園市中壢區", products: [
          { series: "Ritalin LA", yearTarget: 116763, monthly: [12393, 4664, 21477, 16905, 12283, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 107981, monthly: [5472, 3282, 26984, 1046, 12371, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 0, monthly: [0, 1393, 1393, 2752, 2752, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 4180, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 647, 0, 3238, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 723, 0, 1301, 723, null, null, null, null, null, null, null] }
      ] },
      { name: "舒欣身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市平鎮區", products: [
          { series: "Ritalin LA", yearTarget: 155702, monthly: [11986, 16798, 4888, 3639, 14636, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 37018, monthly: [6803, 0, 0, 5442, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 15528, monthly: [6000, 5047, 1905, 4953, 1905, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2880, 7200, 1440, 1440, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5580, monthly: [1238, 0, 0, 1238, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 1222, 0, 0, 1221, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [514, 0, 514, 514, 514, null, null, null, null, null, null, null] }
      ] },
      { name: "天乙藥局(鄭鈞源診所)", channel: "診所(含門前藥局)", system: "", city: "桃園市楊梅區", products: [
          { series: "Lote", yearTarget: 68347, monthly: [6665, 3332, 4998, 8331, 5000, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 58687, monthly: [4632, 3276, 6552, 12285, 6810, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [5473, 8108, 8687, 15318, 2896, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 17485, monthly: [1667, 2838, 0, 1667, 1171, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4458, 0, 743, 743, 2972, null, null, null, null, null, null, null] },
          { series: "Cospirit", yearTarget: 6603, monthly: [0, 5346, 0, 8909, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 3114, 0, 3114, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 3952, monthly: [848, 0, 0, 1504, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3575, monthly: [0, 1429, 0, 3286, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [514, 0, 0, 0, 2324, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 0, monthly: [0, 0, 1306, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "照耀身心精神科診所", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹東鎮", products: [
          { series: "Epine", yearTarget: 45550, monthly: [3286, 3048, 3286, 4428, 3286, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 38905, monthly: [12552, 6276, 6276, 6276, 6276, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 30549, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 25962, monthly: [2731, 2731, 0, 1866, 1867, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 23969, monthly: [8000, 0, 4915, 1372, 3085, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 21598, monthly: [6249, 0, 6250, 2572, 6154, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 2229, 0, 743, 371, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1240, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "和平新竹", channel: "地區醫院", system: "", city: "新竹市北區", products: [
          { series: "Epine", yearTarget: 123040, monthly: [0, 16229, 0, 15847, 0, null, null, null, null, null, null, null] },
          { series: "Cospirit", yearTarget: 31748, monthly: [0, 22552, 0, 22143, 0, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 8009, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [6571, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "均安身心診所", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 92562, monthly: [28254, 0, 7380, 26680, 22825, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3614, 3614, 2169, 3615, 1446, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 12860, monthly: [2571, 0, 3857, 3857, 3858, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2933, monthly: [1257, 0, 0, 1257, 3904, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 2280, monthly: [2286, 0, 1143, 3428, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 1048, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "承美身心診所", channel: "診所(含門前藥局)", system: "", city: "苗栗縣頭份市", products: [
          { series: "Ezole", yearTarget: 70268, monthly: [0, 0, 17524, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 40290, monthly: [17457, 0, 0, 0, 20143, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 21215, monthly: [1367, 0, 0, 3552, 2290, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 9990, monthly: [3333, 0, 3333, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 223, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "瑞安診所", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹東鎮", products: [
          { series: "Epine", yearTarget: 130095, monthly: [0, 0, 6229, 13958, 3571, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 16248, monthly: [0, 0, 0, 0, 4533, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 12136, monthly: [0, 0, 2546, 8486, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 966, 743, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "禾心診所", channel: "診所(含門前藥局)", system: "", city: "苗栗縣頭份市", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [3641, 2913, 10925, 5827, 44876, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 0, monthly: [8727, 11638, 4364, 0, 8623, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [10904, 2181, 4362, 7633, 6543, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [7876, 7084, 4161, 5353, 5352, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [4038, 1010, 1515, 2429, 3399, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [572, 570, 5358, 1786, 714, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 0, monthly: [5585, 2094, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 1381, 0, 0, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 0, monthly: [362, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "龍興診所", channel: "診所(含門前藥局)", system: "", city: "桃園市中壢區", products: [
          { series: "Tone", yearTarget: 61560, monthly: [0, 0, 19428, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 33517, monthly: [0, 15190, 0, 0, 15190, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 11057, monthly: [0, 0, 0, 0, 3086, null, null, null, null, null, null, null] }
      ] },
      { name: "新仁醫院新竹", channel: "地區醫院", system: "", city: "新竹縣竹北市", products: [
          { series: "Exelon Patch", yearTarget: 105305, monthly: [0, 15486, 23229, 0, 14932, null, null, null, null, null, null, null] }
      ] },
      { name: "敏康藥局(敏昌診所)", channel: "診所(含門前藥局)", system: "", city: "桃園市中壢區", products: [
          { series: "Lote", yearTarget: 87230, monthly: [2684, 9393, 4025, 10734, 4025, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [6705, 0, 3353, 0, 3353, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2933, monthly: [838, 838, 1257, 838, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "順心診所(中壢)", channel: "診所(含門前藥局)", system: "", city: "桃園市中壢區", products: [
          { series: "Lote", yearTarget: 71706, monthly: [0, 6290, 0, 7549, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 9350, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1971, 1314, 0, 1314, 2628, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 7065, monthly: [0, 1572, 0, 1572, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新國民醫院", channel: "地區醫院", system: "北醫", city: "桃園市中壢區", products: [
          { series: "Exelon", yearTarget: 56490, monthly: [4468, 0, 2553, 6113, 10950, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [5028, 0, 3353, 3353, 6705, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1696, 2119, 0, 1696, 2543, null, null, null, null, null, null, null] }
      ] },
      { name: "幼幼藥局(六竹診所)", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 45310, monthly: [8092, 0, 0, 8092, 16184, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 8806, monthly: [2516, 2516, 0, 2516, 2516, null, null, null, null, null, null, null] }
      ] },
      { name: "聖母診所", channel: "診所(含門前藥局)", system: "耕莘", city: "新竹縣湖口鄉", products: [
          { series: "Abimay", yearTarget: 21336, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 12010, monthly: [6005, 0, 0, 0, 6005, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [4514, 0, 4514, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1429, 0, 714, 357, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1334, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "楊延壽診所", channel: "診所(含門前藥局)", system: "", city: "桃園市龍潭區", products: [
          { series: "Ritalin LA", yearTarget: 33682, monthly: [0, 0, 2351, 4701, 4701, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1140, 0, 608, 760, null, null, null, null, null, null, null] }
      ] },
      { name: "春天藥局(李傳盛診所)", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹東鎮", products: [
          { series: "Epine", yearTarget: 29060, monthly: [5409, 3429, 0, 0, 3429, null, null, null, null, null, null, null] }
      ] },
      { name: "新中興醫院", channel: "地區醫院", system: "", city: "新竹市北區", products: [
          { series: "Epine", yearTarget: 19840, monthly: [0, 3714, 6190, 0, 7429, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1486, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂地藥局(六竹診所)", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 21037, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 13838, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "夏天藥局(徐鈞平診所)", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹東鎮", products: [
          { series: "Epine", yearTarget: 14770, monthly: [0, 2134, 0, 3143, 0, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 0, monthly: [0, 4114, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安康藥局(億安診所)", channel: "診所(含門前藥局)", system: "", city: "新竹市東區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 13430, 3223, 4297, null, null, null, null, null, null, null] }
      ] },
      { name: "哈哈親子藥局(哈哈親子診所)", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 17496, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "天一診所", channel: "診所(含門前藥局)", system: "", city: "桃園市平鎮區", products: [
          { series: "Fute", yearTarget: 8580, monthly: [0, 0, 1429, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 2660, monthly: [0, 0, 1333, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "崇恩診所", channel: "診所(含門前藥局)", system: "", city: "桃園市中壢區", products: [
          { series: "Tone", yearTarget: 10965, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "佳華藥局(同心海華診所)", channel: "診所(含門前藥局)", system: "", city: "桃園市中壢區", products: [
          { series: "Tone", yearTarget: 4050, monthly: [3239, 0, 0, 3239, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "邱信雄診所", channel: "診所(含門前藥局)", system: "", city: "桃園市中壢區", products: [
          { series: "Ritalin LA", yearTarget: 4619, monthly: [0, 0, 0, 0, 4618, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 432, null, null, null, null, null, null, null] }
      ] },
      { name: "佳美藥局(佑慎診所)", channel: "診所(含門前藥局)", system: "安慎", city: "桃園市中壢區", products: [
          { series: "Exelon", yearTarget: 8442, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "蕭景欽診所", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹東鎮", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 7162, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "嘉岳藥局(翁佩魁診所)", channel: "診所(含門前藥局)", system: "", city: "新竹縣竹北市", products: [
          { series: "Ritalin LA", yearTarget: 2914, monthly: [2914, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "中美醫院", channel: "地區醫院", system: "", city: "桃園市中壢區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [1520, 1520, 760, 760, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 329, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "迎旭診所", channel: "診所(含門前藥局)", system: "", city: "桃園市中壢區", products: [
          { series: "Fute", yearTarget: 0, monthly: [4047, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大安醫院", channel: "地區醫院", system: "", city: "新竹縣竹北市", products: [
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 0, 0, 3581, null, null, null, null, null, null, null] }
      ] },
      { name: "李俊明診所", channel: "診所(含門前藥局)", system: "", city: "桃園市平鎮區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 0, 3181, null, null, null, null, null, null, null] }
      ] },
      { name: "康科特-華揚醫院", channel: "地區醫院", system: "康科特", city: "桃園市中壢區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [1183, 1183, 0, 591, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "滕育成", group: "Div3", customers: [
      { name: "慈惠高雄", channel: "區域醫院", system: "", city: "高雄市大寮區", products: [
          { series: "Ritalin LA", yearTarget: 1740618, monthly: [68628, 113159, 173754, 70812, 123023, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 717615, monthly: [58658, 47699, 61829, 40410, 50007, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 245421, monthly: [21747, 0, 37280, 0, 20127, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 242213, monthly: [30733, 44390, 18439, 30731, 23903, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 206040, monthly: [19047, 6171, 23619, 2857, 12461, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 180790, monthly: [0, 13772, 15762, 0, 13772, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 178167, monthly: [24214, 20080, 35200, 19093, 20944, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 155775, monthly: [7703, 7703, 0, 15406, 5777, null, null, null, null, null, null, null] },
          { series: "Cospirit", yearTarget: 92438, monthly: [17897, 22371, 44743, 19389, 13409, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 86515, monthly: [4286, 4286, 7143, 7143, 5000, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [10000, 11429, 15714, 3571, 4286, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 22715, monthly: [0, 3785, 0, 0, 3785, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [22467, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Suculin", yearTarget: 18659, monthly: [6733, 0, -70, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 7909, monthly: [0, 3906, 2604, 2603, -96, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 2667, 1334, 1334, 1334, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 5936, monthly: [0, 2600, 0, 1857, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "枋寮醫院", channel: "地區醫院", system: "", city: "屏東縣枋寮鄉", products: [
          { series: "Exelon", yearTarget: 847130, monthly: [69872, 88928, 69872, 38112, 69872, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 80962, monthly: [0, 6980, 20940, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 43576, monthly: [8381, 0, 4190, 0, 4119, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 2884, 7214, 21639, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [2077, 1484, 0, 0, 1484, null, null, null, null, null, null, null] }
      ] },
      { name: "屏安醫院", channel: "地區醫院", system: "屏安", city: "屏東縣長治鄉", products: [
          { series: "Abimay", yearTarget: 395170, monthly: [0, 65000, 26000, 0, 25933, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 206930, monthly: [0, 8838, 15952, 12524, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 191562, monthly: [0, 13683, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 146304, monthly: [0, 12952, 0, 15715, 6286, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 27027, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 13440, monthly: [0, 0, 5760, 0, 9635, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3643, 0, 0, 1457, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "快樂心靈診所", channel: "診所(含門前藥局)", system: "快樂心靈", city: "高雄市鳳山區", products: [
          { series: "Ritalin LA", yearTarget: 393412, monthly: [145242, 0, 0, 91881, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 135850, monthly: [21809, 0, 11809, 10000, 11619, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 101802, monthly: [0, 0, 0, 43120, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 87152, monthly: [0, 12571, 0, 0, 12571, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 33670, monthly: [0, 0, 48095, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 21686, 0, 1446, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [6873, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 3114, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "正得藥局(正得身心診所)", channel: "診所(含門前藥局)", system: "正得", city: "屏東縣屏東市", products: [
          { series: "Lote", yearTarget: 387780, monthly: [67428, 33714, 33714, 33714, 33714, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 176799, monthly: [34106, 0, 19249, 23641, 34106, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 55180, monthly: [4429, 4429, 4429, 4428, 4429, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 10388, monthly: [3120, 0, 4159, 4106, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 6900, monthly: [0, 1381, 0, 0, 1381, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 6435, monthly: [0, 0, 1524, 1429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "陳乃菁診所", channel: "診所(含門前藥局)", system: "", city: "高雄市左營區", products: [
          { series: "Exelon", yearTarget: 478144, monthly: [45743, 50667, 28148, 35186, 42223, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 115963, monthly: [4463, 4463, 7439, 17530, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 50220, monthly: [1857, 3715, 3714, 3714, 1857, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 24461, monthly: [3494, 3494, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 13585, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 3352, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 3219, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [363, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "靜安診所", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Lote", yearTarget: 307034, monthly: [50971, 0, 50972, 63714, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 201253, monthly: [20857, 14533, 15895, 17257, 14429, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 13075, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 8092, monthly: [0, 1429, 953, 0, 0, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 3599, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 726, null, null, null, null, null, null, null] }
      ] },
      { name: "元和雅診所", channel: "診所(含門前藥局)", system: "", city: "高雄市鼓山區", products: [
          { series: "Ritalin LA", yearTarget: 270876, monthly: [23077, 25266, 11861, 31715, 41388, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [80571, 53714, 53714, 35922, 3357, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 60990, monthly: [0, 12857, 0, 0, 4257, null, null, null, null, null, null, null] }
      ] },
      { name: "寶建醫院", channel: "區域醫院", system: "", city: "屏東縣屏東市", products: [
          { series: "Exelon", yearTarget: 434056, monthly: [46901, 27868, 13594, 55737, 55058, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 35200, monthly: [2115, 3524, 2467, 2467, 2467, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 11550, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 6812, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1503, 0, 2254, 1503, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3819, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "仁武深忻診所", channel: "診所(含門前藥局)", system: "國良", city: "高雄市仁武區", products: [
          { series: "Ritalin LA", yearTarget: 357957, monthly: [99657, 0, 17833, 45634, 13521, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 71750, monthly: [0, 8191, 0, 0, 8191, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [22500, 0, 4500, 3750, 1500, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 10953, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "欣悅藥局(心悅診所)", channel: "診所(含門前藥局)", system: "心悅", city: "高雄市三民區", products: [
          { series: "Ritalin LA", yearTarget: 177829, monthly: [38572, 0, 14203, 0, 30857, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 71370, monthly: [0, 8353, 0, 11790, 12500, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 59874, monthly: [6667, 0, 9174, 12146, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 48442, monthly: [9144, 0, 6857, 0, 5486, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 48280, monthly: [2723, 0, 8170, 4085, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 46125, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 8857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 0, 1429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "澄清文鳳診所", channel: "診所(含門前藥局)", system: "文鳳", city: "高雄市三民區", products: [
          { series: "Ritalin LA", yearTarget: 319067, monthly: [29158, 29158, 24583, 46499, 37351, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 26910, monthly: [4143, 4143, 0, 4133, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7600, 7600, 3040, 4788, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "五甲心靈診所", channel: "診所(含門前藥局)", system: "快樂心靈", city: "高雄市鳳山區", products: [
          { series: "Ritalin LA", yearTarget: 221606, monthly: [78170, 0, 0, 34286, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 68225, monthly: [0, 0, 8857, 0, 10000, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14458, 14457, 0, 0, 1446, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [5028, 0, 15086, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 14665, monthly: [12571, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 10907, monthly: [3120, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "靜和燕巢", channel: "地區醫院", system: "靜和", city: "高雄市燕巢區", products: [
          { series: "Exprexa", yearTarget: 223391, monthly: [20800, 12800, 10400, 10133, 32800, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 71645, monthly: [31143, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 23008, monthly: [0, 0, 3834, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [14619, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "杏和醫院鳳山", channel: "地區醫院", system: "杏和鳳山", city: "高雄市鳳山區", products: [
          { series: "Exelon", yearTarget: 146183, monthly: [17280, 6912, 10367, 13824, 17280, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 146179, monthly: [16000, 0, 16000, 16000, 16000, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 23100, monthly: [3143, 0, 0, 3143, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 1810, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大社藥局(國良診所)", channel: "診所(含門前藥局)", system: "國良", city: "高雄市路竹區", products: [
          { series: "Ritalin LA", yearTarget: 170352, monthly: [38029, 0, 0, 76057, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 16381, 16381, 20476, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 21216, monthly: [12476, 10953, 21905, 0, 10953, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [11250, 0, 0, 0, 750, null, null, null, null, null, null, null] }
      ] },
      { name: "高安診所", channel: "診所(含門前藥局)", system: "", city: "高雄市新興區", products: [
          { series: "Exprexa", yearTarget: 168286, monthly: [12480, 12480, 12480, 12160, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 108652, monthly: [23620, 0, 0, 0, 23620, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5080, 3629, 2903, 2903, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [3258, 0, 0, 2714, 3257, null, null, null, null, null, null, null] }
      ] },
      { name: "利思達藥局(佳璋診所)", channel: "診所(含門前藥局)", system: "佳璋", city: "高雄市苓雅區", products: [
          { series: "Lote", yearTarget: 116064, monthly: [6238, 6238, 6238, 6238, 6238, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 80750, monthly: [9524, 0, 9524, 9524, 9524, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 26996, monthly: [0, 0, 0, 3857, 7715, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 19812, monthly: [0, 3809, 0, 3619, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7142, 3571, 2143, 2143, 714, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3105, 0, 3105, 3105, 6210, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 7865, monthly: [0, 0, 2857, 0, 1429, null, null, null, null, null, null, null] }
      ] },
      { name: "武廟台安藥局(希望心靈診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Ritalin LA", yearTarget: 164961, monthly: [14854, 16089, 21999, 7449, 19515, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 29250, monthly: [0, 9371, 0, 0, 5857, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 20272, monthly: [0, 0, 0, 2533, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 20148, monthly: [4381, 0, 0, 2457, 2457, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 7228, 0, 0, 1446, null, null, null, null, null, null, null] }
      ] },
      { name: "心寬福藥局(大順景福診所)", channel: "診所(含門前藥局)", system: "寬福", city: "高雄市三民區", products: [
          { series: "Epine", yearTarget: 198033, monthly: [21904, 0, 0, 0, 16761, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 0, 28038, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 21022, monthly: [4572, 0, 0, 0, 4572, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 15914, monthly: [0, 0, 0, 6060, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 8404, monthly: [0, 1866, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 3686, 0, 1474, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 1429, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1390, monthly: [1391, 0, 0, 0, 1391, null, null, null, null, null, null, null] }
      ] },
      { name: "興中好藥局(陳三能診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Ezole", yearTarget: 68328, monthly: [8762, 0, 0, 7619, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 51744, monthly: [0, 0, 9240, 13860, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 49600, monthly: [15525, 0, 0, 6856, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 29330, monthly: [7040, 0, 0, 4481, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [9214, 0, 9214, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7429, 0, 2971, 743, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [4857, 0, 0, 4857, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 4935, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "美心藥局(杏福身心診所)", channel: "診所(含門前藥局)", system: "", city: "屏東縣屏東市", products: [
          { series: "Lote", yearTarget: 81708, monthly: [12381, 2476, 12381, 12381, 12381, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 46791, monthly: [3086, 6171, 12593, 10073, 10074, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 9995, monthly: [0, 2771, 0, 1524, 3048, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 7423, monthly: [4572, 3429, 3429, 5619, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1430, monthly: [0, 0, 0, 1429, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1430, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "和成藥局(和成診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市鳳山區", products: [
          { series: "Exelon", yearTarget: 97932, monthly: [8988, 19355, 34560, 12440, 20737, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 26722, monthly: [0, 1459, 1510, 5670, 11445, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [5855, 0, 6506, 0, 3253, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [2284, 1142, 2427, 1904, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1089, 363, 363, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "精智藥局(佳欣診所)", channel: "診所(含門前藥局)", system: "佳璋", city: "高雄市前鎮區", products: [
          { series: "Lote", yearTarget: 71136, monthly: [3744, 6239, 8733, 6238, 8732, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 40108, monthly: [0, 2315, 3858, 2315, 2315, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 35150, monthly: [1905, 3809, 3810, 5714, 3810, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 21717, monthly: [1142, 1524, 762, 2534, 2534, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 3214, 2857, 1429, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 1242, 1552, 1242, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1430, monthly: [0, 0, 0, 0, 1429, null, null, null, null, null, null, null] }
      ] },
      { name: "季宏診所", channel: "診所(含門前藥局)", system: "", city: "高雄市鳳山區", products: [
          { series: "Ritalin LA", yearTarget: 162154, monthly: [0, 0, 36200, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 16920, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 10843, 1446, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 11600, monthly: [0, 0, 6000, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 8640, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "舒心診所", channel: "診所(含門前藥局)", system: "正得", city: "高雄市三民區", products: [
          { series: "Lote", yearTarget: 89920, monthly: [0, 11238, 0, 11238, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 54224, monthly: [4457, 0, 4457, 0, 4457, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 50730, monthly: [2657, 2657, 2657, 5314, 2657, null, null, null, null, null, null, null] }
      ] },
      { name: "得孚社區藥局(尋路身心診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市左營區", products: [
          { series: "Ezole", yearTarget: 51525, monthly: [4381, 0, 10019, 0, 4381, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 51079, monthly: [6132, 4824, 3336, 5124, 10515, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 41077, monthly: [6586, 0, 0, 0, 3159, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [6531, 1451, 4354, 2177, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 4655, monthly: [7429, 0, 0, 0, 2381, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 0, monthly: [0, 0, 1814, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "文鳳診所", channel: "診所(含門前藥局)", system: "文鳳", city: "高雄市鳳山區", products: [
          { series: "Ritalin LA", yearTarget: 91762, monthly: [16388, 0, 16388, 0, 16388, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 66240, monthly: [20714, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1520, 1520, null, null, null, null, null, null, null] }
      ] },
      { name: "桉家藥師藥局(家欣診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市三民區", products: [
          { series: "Fute", yearTarget: 49335, monthly: [14286, 0, 0, 14286, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 32928, monthly: [20572, 0, 20572, 30857, 30857, null, null, null, null, null, null, null] }
      ] },
      { name: "冠徹藥局(得心身醫學診所)", channel: "診所(含門前藥局)", system: "得立", city: "高雄市左營區", products: [
          { series: "Lote", yearTarget: 68824, monthly: [2457, 7372, 4914, 3685, 3686, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 27767, monthly: [6172, 772, 3085, 5400, 3857, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 19576, monthly: [0, 0, 3066, 2190, 2190, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 13800, monthly: [0, 2771, 924, 1849, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 7253, monthly: [0, 0, 1320, 1306, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 5005, monthly: [0, 1429, 1524, 0, 1429, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3525, monthly: [0, 0, 0, 1410, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "有悅藥局(和誼診所)", channel: "診所(含門前藥局)", system: "佳璋", city: "高雄市新興區", products: [
          { series: "Ritalin LA", yearTarget: 45507, monthly: [23143, 0, 7715, 7715, 15430, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 35150, monthly: [5714, 4762, 4762, 4762, 4762, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [10714, 0, 7143, 3571, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 6240, monthly: [0, 0, 0, 0, 1247, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 5334, monthly: [2286, 1904, 1905, 1810, 1810, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3575, monthly: [0, 1429, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "五甲天民泌尿科診所", channel: "診所(含門前藥局)", system: "天民", city: "高雄市鳳山區", products: [
          { series: "Tone", yearTarget: 127680, monthly: [15238, 0, 15238, 15238, 15238, null, null, null, null, null, null, null] }
      ] },
      { name: "榮欣身心診所", channel: "診所(含門前藥局)", system: "", city: "高雄市新興區", products: [
          { series: "Ezole", yearTarget: 82400, monthly: [0, 0, 0, 20000, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 24412, monthly: [0, 0, 1041, 8107, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 18071, 0, 0, 723, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 16380, monthly: [0, 0, 3514, 0, 5857, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 0, 0, 2942, null, null, null, null, null, null, null] }
      ] },
      { name: "心方診所", channel: "診所(含門前藥局)", system: "", city: "高雄市左營區", products: [
          { series: "Epine", yearTarget: 66030, monthly: [3667, 0, 10977, 0, 10962, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 34000, monthly: [0, 0, 0, 20000, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7600, 0, 15200, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [4571, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 3062, monthly: [4375, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "仁華診所", channel: "診所(含門前藥局)", system: "國良", city: "高雄市仁武區", products: [
          { series: "Ritalin LA", yearTarget: 85176, monthly: [0, 15211, 0, 0, 15211, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 29930, monthly: [16381, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 2625, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "天民泌尿科診所", channel: "診所(含門前藥局)", system: "天民", city: "高雄市三民區", products: [
          { series: "Tone", yearTarget: 110960, monthly: [15238, 0, 15238, 0, 15238, null, null, null, null, null, null, null] }
      ] },
      { name: "養全診所", channel: "診所(含門前藥局)", system: "", city: "高雄市左營區", products: [
          { series: "Lote", yearTarget: 79065, monthly: [6276, 6276, 6276, 6276, 12552, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 10765, monthly: [0, 0, 0, 0, 1538, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 5148, monthly: [4286, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 686, 0, 686, null, null, null, null, null, null, null] }
      ] },
      { name: "惠春藥局(童春診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Exprexa", yearTarget: 41695, monthly: [4774, 0, 5920, 0, 4706, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 31204, monthly: [4000, 0, 4000, 8000, 8000, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5805, monthly: [0, 1286, 0, 0, 1286, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 3686, 0, 737, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 0, 1333, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 981, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "政大聯合藥局(李全忠診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市左營區", products: [
          { series: "Exelon", yearTarget: 64570, monthly: [10705, 0, 4015, 8698, 10037, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 18089, monthly: [3051, 0, 0, 4389, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1657, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 760, null, null, null, null, null, null, null] }
      ] },
      { name: "文化身心診所", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Ezole", yearTarget: 81030, monthly: [17524, 0, 17524, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1503, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "楊明仁診所", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Cospirit", yearTarget: 53915, monthly: [0, 0, 0, 10186, 20371, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 18370, monthly: [3333, 0, 0, 3333, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3907, 0, 0, 0, 376, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 1619, 0, 1619, null, null, null, null, null, null, null] }
      ] },
      { name: "工協身心診所", channel: "診所(含門前藥局)", system: "", city: "高雄市鳳山區", products: [
          { series: "Lote", yearTarget: 22410, monthly: [6223, 0, 0, 2489, 1245, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 18511, monthly: [6943, 0, 1543, 2315, 3857, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7257, 0, 3410, 726, 726, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 8538, monthly: [1172, 0, 1562, 1952, 1171, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 6435, monthly: [0, 0, 0, 1429, 1429, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5520, monthly: [924, 0, 924, 924, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 1201, monthly: [0, 0, 1200, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "輕軌公園藥局(冬勝診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Ezole", yearTarget: 31980, monthly: [2733, 2734, 2343, 2733, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 24320, monthly: [0, 3048, 1524, 3048, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 15953, monthly: [0, 0, 0, 3379, 3379, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 7350, monthly: [0, 0, 1048, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 143, 357, 357, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "河堤診所", channel: "診所(含門前藥局)", system: "寬福", city: "高雄市左營區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [70096, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 12796, monthly: [0, 0, 2743, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7371, 0, 3686, 0, 1474, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5560, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "建工心喜診所", channel: "診所(含門前藥局)", system: "國良", city: "高雄市三民區", products: [
          { series: "Ezole", yearTarget: 30750, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 29660, monthly: [0, 0, 11409, 16733, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 7666, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4351, 0, 1500, 1500, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "寬福診所", channel: "診所(含門前藥局)", system: "寬福", city: "高雄市左營區", products: [
          { series: "Ritalin LA", yearTarget: 42437, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 23764, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7371, 0, 0, 2211, 1474, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3504, 0, 0, 7010, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 8404, monthly: [1867, 0, 0, 1840, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1390, monthly: [0, 0, 0, 1391, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "杏芳藥局(小太陽兒科聯合診所)", channel: "診所(含門前藥局)", system: "", city: "屏東縣屏東市", products: [
          { series: "Ritalin LA", yearTarget: 61776, monthly: [7920, 0, 7920, 7920, 15840, null, null, null, null, null, null, null] }
      ] },
      { name: "鳳山李嘉文泌尿科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市鳳山區", products: [
          { series: "Tone", yearTarget: 37180, monthly: [0, 16857, 0, 15429, 14000, null, null, null, null, null, null, null] }
      ] },
      { name: "健泰診所高雄", channel: "診所(含門前藥局)", system: "", city: "高雄市三民區", products: [
          { series: "Ritalin LA", yearTarget: 22559, monthly: [0, 0, 0, 4667, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 17332, monthly: [2476, 0, 2476, 2476, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 14571, 0, 0, 1457, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 7650, monthly: [1528, 0, 1529, 0, 3047, null, null, null, null, null, null, null] }
      ] },
      { name: "屏安-瑞光藥局", channel: "診所(含門前藥局)", system: "屏安", city: "屏東縣長治鄉", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 2880, 14400, 25252, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7286, 7286, 7286, 3643, 729, null, null, null, null, null, null, null] }
      ] },
      { name: "家慈診所", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Ritalin LA", yearTarget: 51677, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 7143, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "耕心療癒診所", channel: "診所(含門前藥局)", system: "", city: "高雄市左營區", products: [
          { series: "Ritalin LA", yearTarget: 18175, monthly: [3719, 3747, 7437, 0, 3719, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 6058, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 3629, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 3066, monthly: [2629, 0, 2190, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 1800, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "文學藥局(暘明診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市左營區", products: [
          { series: "Abimay", yearTarget: 32117, monthly: [0, 7136, 0, 0, 3567, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 8760, monthly: [0, 2629, 0, 0, 1752, null, null, null, null, null, null, null] }
      ] },
      { name: "文心診所高雄", channel: "診所(含門前藥局)", system: "", city: "高雄市三民區", products: [
          { series: "Ritalin LA", yearTarget: 33264, monthly: [4752, 0, 3167, 3167, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 6900, monthly: [0, 0, 0, 0, 2300, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 1452, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "耀勝藥局(佳和診所)", channel: "診所(含門前藥局)", system: "", city: "屏東縣佳冬鄉", products: [
          { series: "Exelon", yearTarget: 25970, monthly: [0, 0, 2810, 3513, 5621, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5340, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 0, monthly: [0, 0, 0, 2753, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [645, 0, 0, 1934, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "七賢脊椎外科醫院", channel: "地區醫院", system: "七賢", city: "高雄市新興區", products: [
          { series: "Suculin", yearTarget: 47824, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "芯禾診所", channel: "診所(含門前藥局)", system: "", city: "屏東縣屏東市", products: [
          { series: "Exelon", yearTarget: 33328, monthly: [2762, 0, 2762, 4143, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 297, 591, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "康泰診所", channel: "診所(含門前藥局)", system: "", city: "高雄市小港區", products: [
          { series: "Suculin", yearTarget: 27048, monthly: [6187, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3162, 0, 3162, 0, 3162, null, null, null, null, null, null, null] }
      ] },
      { name: "心樂保健藥局(心樂是診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市三民區", products: [
          { series: "Epine", yearTarget: 29640, monthly: [0, 0, 5714, 0, 5714, null, null, null, null, null, null, null] }
      ] },
      { name: "蘇素環神經科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市新興區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 6390, 0, 9586, 6390, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 11760, monthly: [0, 0, 0, 3200, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 743, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "福安診所(屏東)", channel: "診所(含門前藥局)", system: "", city: "屏東縣車城鄉", products: [
          { series: "Epine", yearTarget: 7690, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 4016, monthly: [0, 6519, 5216, 3912, 8475, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 641, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "琉璃光藥師藥局(琉璃光精神科診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [21686, 0, 7229, 7229, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "振隆診所", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Tone", yearTarget: 22880, monthly: [7143, 0, 0, 4285, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大福診所", channel: "診所(含門前藥局)", system: "", city: "高雄市前鎮區", products: [
          { series: "Epine", yearTarget: 23750, monthly: [0, 0, 0, 7619, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1520, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "必佳藥局(杏元診所)", channel: "診所(含門前藥局)", system: "", city: "屏東縣高樹鄉", products: [
          { series: "Exelon", yearTarget: 30290, monthly: [0, 0, 1311, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 979, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "重仁骨科醫院", channel: "地區醫院", system: "", city: "高雄市前金區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [6323, 6323, 6323, 6323, 6323, null, null, null, null, null, null, null] }
      ] },
      { name: "膚悅心診所", channel: "診所(含門前藥局)", system: "國良", city: "高雄市前鎮區", products: [
          { series: "Ritalin LA", yearTarget: 13210, monthly: [7606, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2476, 3750, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "國仁醫院", channel: "地區醫院", system: "國仁", city: "屏東縣屏東市", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [5915, 4600, 3286, 7229, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3990, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "信恩診所", channel: "診所(含門前藥局)", system: "", city: "高雄市新興區", products: [
          { series: "Ezole", yearTarget: 10530, monthly: [3124, 0, 3124, 0, 3124, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3800, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "泰豐健保藥局(富元健康管理診所)", channel: "診所(含門前藥局)", system: "富容", city: "高雄市苓雅區", products: [
          { series: "Exelon", yearTarget: 14296, monthly: [3481, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3191, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "美術清華診所", channel: "診所(含門前藥局)", system: "", city: "高雄市鼓山區", products: [
          { series: "Tone", yearTarget: 7865, monthly: [0, 0, 0, 0, 1429, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 4536, monthly: [0, 0, 3626, 2294, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 369, 369, 369, null, null, null, null, null, null, null] }
      ] },
      { name: "碩恩九如藥局(三民嘉文泌尿科診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市三民區", products: [
          { series: "Tone", yearTarget: 5005, monthly: [3239, 0, 4666, 3239, 4666, null, null, null, null, null, null, null] }
      ] },
      { name: "魏榮洲神經科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市三民區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [18096, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "鳳山高美泌尿科診所", channel: "診所(含門前藥局)", system: "高美", city: "高雄市鳳山區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [3264, 2400, 5664, 4703, 1920, null, null, null, null, null, null, null] }
      ] },
      { name: "陽光診所高雄", channel: "診所(含門前藥局)", system: "正得", city: "高雄市左營區", products: [
          { series: "Ritalin LA", yearTarget: 12563, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5340, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "茂德診所", channel: "診所(含門前藥局)", system: "", city: "高雄市三民區", products: [
          { series: "Ritalin LA", yearTarget: 10027, monthly: [0, 0, 0, 0, 3085, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 2990, 0, 1495, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "七賢光復診所", channel: "診所(含門前藥局)", system: "七賢", city: "高雄市鳳山區", products: [
          { series: "Suculin", yearTarget: 7515, monthly: [1760, 0, 1760, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 3114, 3115, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "屏安-泰祥診所", channel: "診所(含門前藥局)", system: "屏安", city: "屏東縣長治鄉", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 0, 10852, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3643, 0, 0, 729, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "建成診所", channel: "診所(含門前藥局)", system: "", city: "高雄市三民區", products: [
          { series: "Tone", yearTarget: 9295, monthly: [0, 0, 1429, 2858, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "陳琪銘診所", channel: "診所(含門前藥局)", system: "", city: "屏東縣萬丹鄉", products: [
          { series: "Exelon", yearTarget: 5676, monthly: [0, 1383, 0, 1383, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 4166, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 0, 640, null, null, null, null, null, null, null] }
      ] },
      { name: "七賢高美泌尿科診所", channel: "診所(含門前藥局)", system: "高美", city: "高雄市新興區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 7871, 0, 4607, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "文山身心診所", channel: "診所(含門前藥局)", system: "樂群", city: "高雄市三民區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [11400, 0, 0, 0, 760, null, null, null, null, null, null, null] }
      ] },
      { name: "大高田泌尿科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市苓雅區", products: [
          { series: "Tone", yearTarget: 7865, monthly: [0, 0, 0, 2858, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "霖園醫院", channel: "地區醫院", system: "", city: "高雄市林園區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [1643, 2629, 0, 3286, 2629, null, null, null, null, null, null, null] }
      ] },
      { name: "欣聖診所", channel: "診所(含門前藥局)", system: "", city: "高雄市左營區", products: [
          { series: "Tone", yearTarget: 9880, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "屏安-寬懷藥局", channel: "診所(含門前藥局)", system: "屏安", city: "屏東縣長治鄉", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 4800, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 4371, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "開心藥局(樂群診所)", channel: "診所(含門前藥局)", system: "樂群", city: "高雄市鼓山區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [7600, 0, 0, 1520, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "富容健保藥局(德容骨外科診所)", channel: "診所(含門前藥局)", system: "富容", city: "高雄市三民區", products: [
          { series: "Exelon", yearTarget: 8572, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "華榮骨科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市鼓山區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [4251, 0, 0, 0, 4251, null, null, null, null, null, null, null] }
      ] },
      { name: "柯偉恭診所", channel: "診所(含門前藥局)", system: "", city: "高雄市新興區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 7503, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "得安藥局(得立診所)", channel: "診所(含門前藥局)", system: "", city: "屏東縣屏東市", products: [
          { series: "Lote", yearTarget: 0, monthly: [6223, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "秦文鎮診所", channel: "診所(含門前藥局)", system: "", city: "屏東縣東港鎮", products: [
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 0, 0, 6095, null, null, null, null, null, null, null] }
      ] },
      { name: "前進診所", channel: "診所(含門前藥局)", system: "", city: "屏東縣屏東市", products: [
          { series: "Exelon", yearTarget: 5871, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "信宜診所", channel: "診所(含門前藥局)", system: "", city: "高雄市鳳山區", products: [
          { series: "Epine", yearTarget: 1690, monthly: [3381, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "達仁診所", channel: "診所(含門前藥局)", system: "", city: "高雄市前鎮區", products: [
          { series: "Exelon Patch", yearTarget: 0, monthly: [4398, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大興骨科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市三民區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [4251, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "祐新骨外科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市左營區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 0, 4251, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "聖恩診所", channel: "診所(含門前藥局)", system: "杏和鳳山", city: "高雄市鳳山區", products: [
          { series: "Slatone", yearTarget: 4200, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "屏安-心安藥局(興安診所)", channel: "診所(含門前藥局)", system: "屏安", city: "屏東縣屏東市", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 0, 3617, null, null, null, null, null, null, null] }
      ] },
      { name: "宸心藥局(常紳診所)", channel: "診所(含門前藥局)", system: "", city: "屏東縣屏東市", products: [
          { series: "Tone", yearTarget: 3575, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "張美玉泌尿科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市小港區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 3380, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "湖內健康藥局(健康湖內診所)", channel: "診所(含門前藥局)", system: "富容", city: "高雄市湖內區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [3191, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "誠心診所", channel: "診所(含門前藥局)", system: "", city: "高雄市新興區", products: [
          { series: "Exelon Patch", yearTarget: 0, monthly: [0, 0, 0, 0, 2885, null, null, null, null, null, null, null] }
      ] },
      { name: "聯合王婦產科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市林園區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 0, 2714, null, null, null, null, null, null, null] }
      ] },
      { name: "雄大診所", channel: "診所(含門前藥局)", system: "", city: "高雄市新興區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 2714, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "柏仁醫院", channel: "地區醫院", system: "", city: "高雄市左營區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [363, 0, 435, 363, 581, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "劉進輝", group: "Div3", customers: [
      { name: "樂行診所", channel: "診所(含門前藥局)", system: "", city: "台南市仁德區", products: [
          { series: "Ritalin LA", yearTarget: 1156782, monthly: [254571, 194143, 0, 257400, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 41143, 0, 0, 4800, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 18240, monthly: [0, 3048, 0, 4571, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 10514, 7010, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "舒心藥局(心寬診所)", channel: "診所(含門前藥局)", system: "心寬", city: "台南市永康區", products: [
          { series: "Ritalin LA", yearTarget: 865777, monthly: [80574, 60657, 75646, 99828, 44965, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 178405, monthly: [22143, 0, 0, 21071, 21071, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 166926, monthly: [31571, 0, 12429, 31571, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 87237, monthly: [21066, 0, 0, 9067, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [20572, 0, 27429, 13714, 2057, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 9690, monthly: [3429, 0, 0, 3429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "春暉診所", channel: "診所(含門前藥局)", system: "", city: "台南市東區", products: [
          { series: "Ritalin LA", yearTarget: 1032035, monthly: [268114, 0, 95040, 160348, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [45600, 45600, 22800, 22800, 7600, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 102079, monthly: [12685, 9296, 9514, 6971, 12210, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 10125, monthly: [1353, 0, 1353, 0, 2704, null, null, null, null, null, null, null] }
      ] },
      { name: "心悅藥局(安平心寬診所)", channel: "診所(含門前藥局)", system: "心寬", city: "台南市安平區", products: [
          { series: "Ritalin LA", yearTarget: 785067, monthly: [75841, 60658, 84412, 84840, 84841, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 153264, monthly: [31571, 8286, 0, 39609, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [13714, 6857, 13714, 7543, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 39800, monthly: [1962, 0, 5905, 0, 9543, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 31449, monthly: [2347, 0, 3520, 0, 5440, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3990, monthly: [0, 2286, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "好藥藥局(吳南逸診所)", channel: "診所(含門前藥局)", system: "", city: "嘉義市西區", products: [
          { series: "Ritalin LA", yearTarget: 314688, monthly: [96886, 0, 0, 133977, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [100000, 0, 69524, 0, 30476, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 163530, monthly: [47429, 0, 0, 47429, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 59520, monthly: [19238, 0, 19238, 0, 19238, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 1900, 2280, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "有誼藥局(誼仁診所)", channel: "診所(含門前藥局)", system: "", city: "台南市善化區", products: [
          { series: "Ritalin LA", yearTarget: 651484, monthly: [165679, 0, 73828, 18023, 73828, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 41667, monthly: [17360, 0, 0, 0, 16857, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [9657, 9657, 0, 11886, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 6345, monthly: [0, 7048, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "陳俊升診所", channel: "診所(含門前藥局)", system: "", city: "台南市新營區", products: [
          { series: "Ritalin LA", yearTarget: 278296, monthly: [24723, 6746, 23594, 24609, 32029, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [34263, 34263, 34263, 0, 34263, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 93726, monthly: [5715, 11429, 0, 6857, 6857, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 79380, monthly: [4048, 4857, 9714, 4048, 9714, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 43680, monthly: [3905, 3905, 3905, 2343, 6248, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 41370, monthly: [0, 2571, 0, 2572, 2572, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [9314, 6210, 9314, 12418, 0, null, null, null, null, null, null, null] },
          { series: "Cospirit", yearTarget: 33699, monthly: [16867, 0, 7667, 15047, 15048, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [10520, 0, 7514, 6011, 751, null, null, null, null, null, null, null] }
      ] },
      { name: "香柏藥局(東橋身心診所)", channel: "診所(含門前藥局)", system: "", city: "台南市永康區", products: [
          { series: "Ritalin LA", yearTarget: 347206, monthly: [18977, 28142, 0, 27850, 37370, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 142660, monthly: [24266, 0, 15039, 14753, 8943, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 125120, monthly: [10895, 0, 10895, 10895, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 39468, monthly: [8571, 8571, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14514, 7257, 7257, 8708, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3514, 0, 0, 0, 3514, null, null, null, null, null, null, null] }
      ] },
      { name: "芳鄰藥局(殷建智精神科診所)", channel: "診所(含門前藥局)", system: "", city: "台南市北區", products: [
          { series: "Ritalin LA", yearTarget: 430473, monthly: [23560, 35340, 36100, 35974, 47120, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 45990, monthly: [0, 8762, 0, 8762, 8762, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [22543, 0, 7514, 7514, 1503, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 7029, 3048, 0, 7029, null, null, null, null, null, null, null] }
      ] },
      { name: "郭玉柱診所", channel: "診所(含門前藥局)", system: "郭玉柱", city: "高雄市楠梓區", products: [
          { series: "Lote", yearTarget: 345000, monthly: [30000, 30000, 60000, 35000, 30000, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 56795, monthly: [0, 3067, 13800, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 37742, monthly: [0, 3932, 4718, 7076, 6291, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 19728, monthly: [2880, 0, 8228, 0, 8000, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 3751, 0, 1501, 750, null, null, null, null, null, null, null] }
      ] },
      { name: "心欣診所", channel: "診所(含門前藥局)", system: "", city: "高雄市楠梓區", products: [
          { series: "Ritalin LA", yearTarget: 330295, monthly: [12651, 55428, 0, 30794, 58794, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 107440, monthly: [13429, 13429, 0, 13429, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5943, 5943, 3714, 4457, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 9384, monthly: [2762, 0, 2762, 0, 2657, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 3600, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂安醫院", channel: "地區醫院", system: "樂安", city: "高雄市岡山區", products: [
          { series: "Ezole", yearTarget: 134316, monthly: [8190, 4095, 8190, 11857, 7905, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 71820, monthly: [8695, 3238, 8695, 5457, 10914, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 67498, monthly: [10800, 10800, 10800, 0, 10347, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [12610, 10087, 12610, 12610, 15762, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 61834, monthly: [0, 7190, 5752, 7190, 14381, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 48097, monthly: [3657, 2194, 4279, 2256, 3610, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 26813, monthly: [1914, 0, 1914, 0, 1800, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3559, 3029, 0, 2271, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [3239, 0, 0, 1619, 3239, null, null, null, null, null, null, null] }
      ] },
      { name: "心自在身心診所", channel: "診所(含門前藥局)", system: "心樂活", city: "台南市東區", products: [
          { series: "Ritalin LA", yearTarget: 380250, monthly: [137006, 68503, 2468, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [22714, 22714, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 4340, monthly: [8666, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 2334, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "王建安小兒科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市楠梓區", products: [
          { series: "Ritalin LA", yearTarget: 354816, monthly: [39600, 0, 39600, 79200, 39600, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7503, 7503, 0, 3752, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "欣悅診所台南", channel: "診所(含門前藥局)", system: "", city: "台南市南區", products: [
          { series: "Ritalin LA", yearTarget: 190637, monthly: [31181, 19565, 11616, 30963, 30963, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 66770, monthly: [0, 0, 0, 29948, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 47292, monthly: [21600, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7514, 0, 7514, 0, 2254, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 6495, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "維佳身心診所", channel: "診所(含門前藥局)", system: "心悅", city: "高雄市楠梓區", products: [
          { series: "Ritalin LA", yearTarget: 119263, monthly: [60489, 14203, 30857, 4735, 14203, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 70720, monthly: [0, 8172, 8172, 6810, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 46075, monthly: [0, 0, 6666, 0, 3240, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 42940, monthly: [3181, 6762, 3181, 0, 3428, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 19194, monthly: [0, 0, 2743, 4571, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 8937, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 8517, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 1429, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "心永康身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "台南市永康區", products: [
          { series: "Ezole", yearTarget: 180180, monthly: [51430, 0, 25715, 24857, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 79804, monthly: [4495, 8991, 0, 6743, 6743, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 37240, monthly: [3048, 3809, 0, 4372, 8543, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 18240, monthly: [2285, 2285, 1143, 4571, 4571, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4900, 0, 3500, 700, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林俞仲身心診所", channel: "診所(含門前藥局)", system: "林俞仲", city: "台南市永康區", products: [
          { series: "Lote", yearTarget: 129835, monthly: [0, 22572, 22572, 22572, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 89110, monthly: [0, 39360, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 52440, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 15269, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 6975, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "明澤欣心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "台南市永康區", products: [
          { series: "Ritalin LA", yearTarget: 133092, monthly: [24434, 7383, 0, 19600, 12217, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 74184, monthly: [0, 0, 11239, 11238, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 25080, monthly: [0, 2286, 0, 3810, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7000, 0, 7000, 0, 1400, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 15015, monthly: [0, 0, 4286, 7620, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 14820, monthly: [0, 0, 3428, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大和藥局(愛仁診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市大樹區", products: [
          { series: "Exelon", yearTarget: 185327, monthly: [22484, 13228, 14550, 36373, 21822, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 30570, monthly: [1546, 0, 3091, 1489, 6058, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 7713, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 6380, monthly: [3476, 1428, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [2687, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "迦南內科診所", channel: "診所(含門前藥局)", system: "", city: "台南市南區", products: [
          { series: "Lote", yearTarget: 107630, monthly: [17171, 17171, 17171, 17171, 17171, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 66490, monthly: [0, 0, 6095, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 54720, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 0, 8229, 12343, null, null, null, null, null, null, null] }
      ] },
      { name: "春霖診所", channel: "診所(含門前藥局)", system: "", city: "台南市北區", products: [
          { series: "Tone", yearTarget: 191100, monthly: [42191, 0, 42190, 0, 42191, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [8024, 0, 9671, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "恩泉藥局(活水神經內科診所)", channel: "診所(含門前藥局)", system: "活水", city: "台南市東區", products: [
          { series: "Exelon", yearTarget: 150141, monthly: [0, 10917, 17464, 16009, 10188, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 67830, monthly: [0, 0, 6666, 4000, 6667, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 15596, 936, 6550, 10293, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 3658, 1829, 1829, null, null, null, null, null, null, null] }
      ] },
      { name: "心悠活診所", channel: "診所(含門前藥局)", system: "心樂活", city: "台南市北區", products: [
          { series: "Ritalin LA", yearTarget: 178811, monthly: [33601, 0, 0, 19029, 25131, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [15143, 0, 0, 3786, 757, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 11780, monthly: [0, 0, 2477, 0, 1238, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 4210, monthly: [905, 0, 0, 2714, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "佳美藥局(安河神經內科診所)", channel: "診所(含門前藥局)", system: "活水", city: "台南市安南區", products: [
          { series: "Exelon", yearTarget: 150892, monthly: [0, 8005, 8733, 6549, 2910, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 67830, monthly: [0, 0, 2666, 5333, 9333, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 12476, 0, 0, 2184, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 1829, 1829, 1829, 3658, null, null, null, null, null, null, null] }
      ] },
      { name: "悅心診所", channel: "診所(含門前藥局)", system: "郭玉柱", city: "高雄市楠梓區", products: [
          { series: "Lote", yearTarget: 132000, monthly: [0, 0, 30000, 0, 50000, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 14385, monthly: [0, 0, 4114, 0, 8000, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7503, 0, 0, 0, 1501, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 6290, monthly: [0, 0, 0, 7863, 15725, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 6140, monthly: [0, 0, 0, 1524, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "利安藥局(知心連冀診所)", channel: "診所(含門前藥局)", system: "", city: "嘉義市東區", products: [
          { series: "Ritalin LA", yearTarget: 153555, monthly: [12208, 7637, 12209, 4572, 7637, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 28960, monthly: [0, 0, 5476, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 28382, monthly: [6171, 0, 6171, 0, 6171, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [4810, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "同行身心診所", channel: "診所(含門前藥局)", system: "", city: "台南市新市區", products: [
          { series: "Ritalin LA", yearTarget: 165867, monthly: [0, 10432, 0, 9326, 11757, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 30212, monthly: [0, 0, 0, 0, 2800, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 17010, monthly: [0, 0, 0, 0, 5671, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1080, 0, 720, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1370, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "思元藥局(康舟診所)", channel: "診所(含門前藥局)", system: "康舟", city: "台南市南區", products: [
          { series: "Ritalin LA", yearTarget: 83853, monthly: [16575, 0, 8115, 4403, 12518, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 63792, monthly: [8857, 0, 8857, 0, 8858, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7257, 0, 7257, 2177, 2902, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [2270, 0, 2270, 0, 508, null, null, null, null, null, null, null] }
      ] },
      { name: "快樂心寧身心診所", channel: "診所(含門前藥局)", system: "", city: "台南市永康區", products: [
          { series: "Ritalin LA", yearTarget: 102869, monthly: [15245, 24517, 21097, 23417, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 2572, 4286, 7904, 9915, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4457, 0, 4457, 2229, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1430, monthly: [0, 0, 0, 1429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂寧藥局(聖星診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市旗山區", products: [
          { series: "Exelon", yearTarget: 111192, monthly: [14346, 25107, 14347, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [11417, 13840, 0, 0, 1300, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 10962, monthly: [0, 2349, 1302, 0, 2604, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [3582, 3580, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2194, 2926, 1463, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 0, 1691, null, null, null, null, null, null, null] }
      ] },
      { name: "祈福診所", channel: "診所(含門前藥局)", system: "", city: "高雄市楠梓區", products: [
          { series: "Ritalin LA", yearTarget: 59942, monthly: [13468, 0, 0, 13468, 13468, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 36432, monthly: [3962, 0, 0, 0, 3905, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 35420, monthly: [7695, 0, 0, 3724, 7448, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 7514, 0, 1503, null, null, null, null, null, null, null] }
      ] },
      { name: "安大身心精神科診所", channel: "診所(含門前藥局)", system: "林俞仲", city: "台南市安南區", products: [
          { series: "Ezole", yearTarget: 112157, monthly: [18743, 0, 0, 0, 11048, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 12419, monthly: [0, 11286, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 10912, monthly: [0, 0, 2834, 0, 2835, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 7410, monthly: [0, 0, 0, 0, 3428, null, null, null, null, null, null, null] }
      ] },
      { name: "笑嘻嘻藥局(蕭尹瑩身心科診所)", channel: "診所(含門前藥局)", system: "蕭文勝", city: "台南市北區", products: [
          { series: "Ezole", yearTarget: 65532, monthly: [11428, 0, 0, 11086, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 39196, monthly: [13209, 0, 0, 19814, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [18714, 0, 0, 2246, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "聖安藥局(陳俊憲神經內科診所)", channel: "診所(含門前藥局)", system: "", city: "台南市南區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [10842, 19714, 20699, 10842, 19714, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 56358, monthly: [0, 2815, 0, 11967, 5631, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 9295, monthly: [1429, 0, 1429, 1429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安馨民雄診所", channel: "診所(含門前藥局)", system: "嘉基", city: "嘉義縣民雄鄉", products: [
          { series: "Exelon", yearTarget: 73470, monthly: [5726, 0, 11450, 0, 8588, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 35235, monthly: [0, 0, 0, 0, 1215, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 7810, monthly: [714, 0, 1429, 1429, 0, null, null, null, null, null, null, null] },
          { series: "Epram", yearTarget: 5394, monthly: [1400, 933, 0, 0, 933, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 226, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "禾元藥局(心禾康舟診所)", channel: "診所(含門前藥局)", system: "康舟", city: "台南市北區", products: [
          { series: "Ritalin LA", yearTarget: 64035, monthly: [29627, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 23922, monthly: [0, 3542, 3542, 4428, 2657, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 6531, 0, 3629, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "幸運藥局(蕭文勝診所)", channel: "診所(含門前藥局)", system: "蕭文勝", city: "台南市東區", products: [
          { series: "Ezole", yearTarget: 79248, monthly: [7619, 7619, 381, 7391, 7391, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 18000, monthly: [0, 0, 3600, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 749, 2246, 0, 749, null, null, null, null, null, null, null] }
      ] },
      { name: "保心藥局(心田診所)", channel: "診所(含門前藥局)", system: "", city: "台南市歸仁區", products: [
          { series: "Ritalin LA", yearTarget: 64701, monthly: [6635, 4976, 0, 0, 14930, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3538, 0, 0, 0, 737, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1520, monthly: [0, 0, 0, 0, 1524, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 1176, monthly: [785, 2354, 0, 0, 2354, null, null, null, null, null, null, null] }
      ] },
      { name: "心樂活診所", channel: "診所(含門前藥局)", system: "心樂活", city: "台南市東區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [22714, 0, 21200, 7571, 2271, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 21844, monthly: [6686, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 4660, monthly: [2714, 4142, 0, 0, 2714, null, null, null, null, null, null, null] }
      ] },
      { name: "陳文勝診所", channel: "診所(含門前藥局)", system: "", city: "嘉義市東區", products: [
          { series: "Epine", yearTarget: 60450, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 15715, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 9730, monthly: [0, 0, 0, 0, 2781, null, null, null, null, null, null, null] }
      ] },
      { name: "仁享診所", channel: "診所(含門前藥局)", system: "", city: "台南市永康區", products: [
          { series: "Lote", yearTarget: 55430, monthly: [0, 12048, 0, 0, 12048, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 2184, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "東祐身心診所", channel: "診所(含門前藥局)", system: "", city: "台南市新營區", products: [
          { series: "Ritalin LA", yearTarget: 25820, monthly: [0, 7377, 0, 14754, 22131, null, null, null, null, null, null, null] }
      ] },
      { name: "北海藥局(曾泌尿科診所)", channel: "診所(含門前藥局)", system: "", city: "台南市北區", products: [
          { series: "Tone", yearTarget: 43860, monthly: [12858, 0, 0, 0, 12858, null, null, null, null, null, null, null] }
      ] },
      { name: "光雄長安醫院", channel: "地區醫院", system: "", city: "高雄市岡山區", products: [
          { series: "Exelon", yearTarget: 35155, monthly: [6460, 9690, 0, 9690, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1264, 1264, 0, 0, 2845, null, null, null, null, null, null, null] }
      ] },
      { name: "江錫輝診所", channel: "診所(含門前藥局)", system: "", city: "台南市東區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 19257, 0, 38515, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "真心診所台南", channel: "診所(含門前藥局)", system: "", city: "台南市永康區", products: [
          { series: "Ritalin LA", yearTarget: 23468, monthly: [4074, 983, 3977, 7703, 4043, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 2205, monthly: [1467, 1467, 0, 2934, 2934, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [743, 0, 0, 0, 371, null, null, null, null, null, null, null] }
      ] },
      { name: "幼安藥局(陽光小兒科診所)", channel: "診所(含門前藥局)", system: "", city: "台南市安平區", products: [
          { series: "Ritalin LA", yearTarget: 27160, monthly: [0, 8165, 4778, 6254, 4017, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 2499, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 581, 363, 363, 363, null, null, null, null, null, null, null] }
      ] },
      { name: "胡崇元活泉診所", channel: "診所(含門前藥局)", system: "", city: "台南市東區", products: [
          { series: "Ezole", yearTarget: 30130, monthly: [6552, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3571, 0, 3571, 0, 3571, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1676, 0, 0, 0, 1676, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [946, 0, 0, 0, 1891, null, null, null, null, null, null, null] }
      ] },
      { name: "北榮藥局(宋思權診所)", channel: "診所(含門前藥局)", system: "", city: "嘉義市西區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [21833, 0, 15596, 0, 15596, null, null, null, null, null, null, null] }
      ] },
      { name: "多樂藥局(林晏弘診所)", channel: "診所(含門前藥局)", system: "", city: "台南市新營區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 13181, 13181, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 22500, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 2404, 0, 751, null, null, null, null, null, null, null] }
      ] },
      { name: "侯守恩診所", channel: "診所(含門前藥局)", system: "", city: "嘉義縣朴子市", products: [
          { series: "Ritalin LA", yearTarget: 23743, monthly: [6849, 0, 6849, 0, 6849, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [754, 0, 1508, 0, 754, null, null, null, null, null, null, null] }
      ] },
      { name: "康福藥師藥局(穠田聯合診所)", channel: "診所(含門前藥局)", system: "", city: "台南市安平區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [9843, 0, 0, 9843, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 18709, monthly: [4374, 0, 0, 0, 4375, null, null, null, null, null, null, null] }
      ] },
      { name: "永和醫院台南", channel: "地區醫院", system: "", city: "台南市中西區", products: [
          { series: "Exelon Patch", yearTarget: 0, monthly: [9087, 0, 7714, 0, 7714, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 11208, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 7140, monthly: [0, 1190, 0, 0, 1191, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 0, 1790, null, null, null, null, null, null, null] }
      ] },
      { name: "周裕軒診所", channel: "診所(含門前藥局)", system: "", city: "嘉義市東區", products: [
          { series: "Ritalin LA", yearTarget: 34096, monthly: [0, 0, 0, 7749, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "明星藥局(明如身心診所)", channel: "診所(含門前藥局)", system: "", city: "台南市佳里區", products: [
          { series: "Exprexa", yearTarget: 30943, monthly: [8640, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂森藥局(許森彥精神科診所)", channel: "診所(含門前藥局)", system: "", city: "台南市北區", products: [
          { series: "Exprexa", yearTarget: 13482, monthly: [0, 0, 4496, 0, 2997, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 8800, monthly: [0, 3286, 0, 0, 3286, null, null, null, null, null, null, null] }
      ] },
      { name: "昱安藥局(潘內兒科)", channel: "診所(含門前藥局)", system: "", city: "嘉義市西區", products: [
          { series: "Epine", yearTarget: 15600, monthly: [3905, 3905, 3905, 3905, 3905, null, null, null, null, null, null, null] }
      ] },
      { name: "欣明精神科診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "高雄市旗山區", products: [
          { series: "Ezole", yearTarget: 19734, monthly: [4286, 0, 0, 4143, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 1400, 1400, 1400, null, null, null, null, null, null, null] }
      ] },
      { name: "甫原中西人合藥局(林邑樵耳鼻喉科診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市楠梓區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [4114, 0, 0, 4114, 4114, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 3138, 4708, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 5231, monthly: [0, 0, 0, 4368, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "維心診所", channel: "診所(含門前藥局)", system: "", city: "高雄市岡山區", products: [
          { series: "Epine", yearTarget: 18280, monthly: [0, 4571, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3757, 0, 751, 751, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1640, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "德泰健保藥局(佳愛診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市大社區", products: [
          { series: "Tone", yearTarget: 15470, monthly: [0, 2715, 10334, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "楠梓臻心診所", channel: "診所(含門前藥局)", system: "", city: "高雄市楠梓區", products: [
          { series: "Ezole", yearTarget: 7645, monthly: [0, 2781, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 4628, monthly: [0, 0, 0, 3857, 7715, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 1080, 0, 720, null, null, null, null, null, null, null] }
      ] },
      { name: "永恒藥局(蔡政達診所)", channel: "診所(含門前藥局)", system: "", city: "高雄市梓官區", products: [
          { series: "Abimay", yearTarget: 5667, monthly: [5667, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 0, 3238, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 3172, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 2953, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 2100, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安芯診所", channel: "診所(含門前藥局)", system: "", city: "台南市麻豆區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [11400, 0, 3800, 3800, 1520, null, null, null, null, null, null, null] }
      ] },
      { name: "欣平藥局(第一聯合診所)", channel: "診所(含門前藥局)", system: "", city: "台南市北區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [7105, 0, 0, 7105, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 0, 1824, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [749, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "麗生藥師藥局(許乃仁皮膚科診所)", channel: "診所(含門前藥局)", system: "", city: "台南市東區", products: [
          { series: "Suculin", yearTarget: 16523, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "祥和診所", channel: "診所(含門前藥局)", system: "", city: "台南市關廟區", products: [
          { series: "Exelon", yearTarget: 15980, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "太子藥局永大店(楊寧娟內科診所)", channel: "診所(含門前藥局)", system: "", city: "台南市永康區", products: [
          { series: "Tone", yearTarget: 8360, monthly: [0, 3239, 3048, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "徐鴻傑身心診所", channel: "診所(含門前藥局)", system: "", city: "嘉義縣民雄鄉", products: [
          { series: "Exprexa", yearTarget: 13627, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "晟欣診所", channel: "診所(含門前藥局)", system: "", city: "台南市中西區", products: [
          { series: "Tone", yearTarget: 7865, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5520, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大金田泌尿科診所", channel: "診所(含門前藥局)", system: "", city: "高雄市楠梓區", products: [
          { series: "Tone", yearTarget: 8645, monthly: [0, 0, 0, 4000, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "楠梓心寬診所", channel: "診所(含門前藥局)", system: "郭玉柱", city: "高雄市楠梓區", products: [
          { series: "Epine", yearTarget: 9210, monthly: [1534, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1200, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂而康診所", channel: "診所(含門前藥局)", system: "", city: "台南市東區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 2376, 3960, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 1248, monthly: [0, 0, 2495, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1114, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "永大診所", channel: "診所(含門前藥局)", system: "", city: "台南市永康區", products: [
          { series: "Fute", yearTarget: 5805, monthly: [0, 2572, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "韓內兒科診所", channel: "診所(含門前藥局)", system: "", city: "台南市中西區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 0, 6248, null, null, null, null, null, null, null] }
      ] },
      { name: "美華健保藥局(陳相國聯合診所)", channel: "診所(含門前藥局)", system: "", city: "台南市新化區", products: [
          { series: "Tone", yearTarget: 1520, monthly: [0, 3048, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "德南診所", channel: "診所(含門前藥局)", system: "", city: "嘉義縣鹿草鄉", products: [
          { series: "Exelon", yearTarget: 3850, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "以琳內科診所", channel: "診所(含門前藥局)", system: "", city: "台南市佳里區", products: [
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 0, 3477, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "陳冠文診所", channel: "診所(含門前藥局)", system: "", city: "台南市北區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [3337, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "長生診所", channel: "診所(含門前藥局)", system: "", city: "嘉義縣東石鄉", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 3286, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "翁桂芳精神科診所", channel: "診所(含門前藥局)", system: "", city: "台南市中西區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 2399, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安安婦幼診所", channel: "診所(含門前藥局)", system: "", city: "台南市北區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1520, 760, null, null, null, null, null, null, null] }
      ] },
      { name: "璟馨婦幼醫院", channel: "地區醫院", system: "", city: "台南市永康區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 1791, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "高大美杏生醫院", channel: "地區醫院", system: "", city: "高雄市楠梓區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 972, 0, 647, null, null, null, null, null, null, null] }
      ] },
      { name: "蘇炳文內科診所", channel: "診所(含門前藥局)", system: "", city: "台南市善化區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [1605, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安馨嘉義內科診所", channel: "診所(含門前藥局)", system: "嘉基", city: "嘉義市西區", products: [
          { series: "Epine", yearTarget: 1420, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "靜心診所", channel: "診所(含門前藥局)", system: "", city: "台南市東區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 371, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "GP空區", group: "Div3", customers: [
      { name: "卓大夫診所", channel: "診所(含門前藥局)", system: "", city: "台中市西屯區", products: [
          { series: "Ritalin LA", yearTarget: 829137, monthly: [18491, 228342, 0, 18491, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 812725, monthly: [82286, 82286, 82286, 82286, 41143, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 565340, monthly: [49143, 0, 49143, 49143, 49143, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 466040, monthly: [69238, 42286, 21143, 42286, 21143, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 374235, monthly: [44572, 0, 44572, 43943, 19429, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 99656, monthly: [14743, 0, 10200, 14203, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [44571, 0, 0, 0, 6687, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 35588, monthly: [2293, 2293, 2293, 4480, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 5130, monthly: [1143, 0, 1143, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "橙易藥局(吳四維診所)", channel: "診所(含門前藥局)", system: "", city: "苗栗縣苗栗市", products: [
          { series: "Ritalin LA", yearTarget: 1030565, monthly: [144450, 29257, 36572, 108689, 213121, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 761540, monthly: [91934, 39400, 78800, 78800, 39400, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 248868, monthly: [0, 26362, 16191, 26148, 26148, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [27658, 27657, 20743, 24200, 4840, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 62045, monthly: [3514, 11371, 0, 3429, 11142, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 39984, monthly: [5333, 5333, 2667, 10346, 5120, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3575, monthly: [5715, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "心靜藥局(趙玉良診所)", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Ritalin LA", yearTarget: 750666, monthly: [19800, 61143, 61143, 80943, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 370890, monthly: [56666, 25714, 17143, 37143, 44762, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 135815, monthly: [23620, 0, 23620, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 45552, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 14286, 7143, 1429, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3575, monthly: [0, 0, 0, 0, 2858, null, null, null, null, null, null, null] }
      ] },
      { name: "堅美身心診所", channel: "診所(含門前藥局)", system: "", city: "台中市西屯區", products: [
          { series: "Ritalin LA", yearTarget: 626067, monthly: [97383, 60811, 47520, 87771, 140011, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 70610, monthly: [13067, 0, 12533, 0, 12533, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 63440, monthly: [0, 0, 9143, 9143, 9143, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 39755, monthly: [4762, 2857, 0, 2857, 4762, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7200, 0, 2160, 6480, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 0, 0, 13095, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 7865, monthly: [0, 0, 2857, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "中科好晴天藥局(中科好晴天診所)", channel: "診所(含門前藥局)", system: "好晴天", city: "台中市大雅區", products: [
          { series: "Ritalin LA", yearTarget: 541990, monthly: [57192, 61801, 61801, 38758, 86520, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 108730, monthly: [8071, 0, 8071, 17214, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 35700, monthly: [0, 0, 0, 5953, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 14286, 7143, 3571, 1428, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 25201, monthly: [3172, 0, 3171, 0, 3086, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5180, monthly: [0, 1476, 0, 1476, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 5005, monthly: [0, 0, 1429, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "文心好晴天藥局(文心好晴天身心診所)", channel: "診所(含門前藥局)", system: "好晴天", city: "台中市南屯區", products: [
          { series: "Ritalin LA", yearTarget: 502020, monthly: [55810, 88992, 0, 88992, 55810, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 51016, monthly: [0, 0, 17214, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 28571, 7143, 0, 714, null, null, null, null, null, null, null] },
          { series: "Suculin", yearTarget: 9863, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 5950, monthly: [0, 0, 0, 0, 1190, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3575, monthly: [0, 1429, 1429, 0, 2858, null, null, null, null, null, null, null] }
      ] },
      { name: "清海醫院", channel: "地區醫院", system: "", city: "台中市石岡區", products: [
          { series: "Fute", yearTarget: 398300, monthly: [0, 80172, 0, 83791, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 92231, monthly: [0, 6190, 6190, 0, 12381, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 83840, monthly: [10476, 10476, 0, 10000, 10000, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 54985, monthly: [0, 15714, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 9698, monthly: [0, 4042, 4042, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 4857, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "永信藥師藥局(天慈身心診所)", channel: "診所(含門前藥局)", system: "", city: "苗栗縣苑裡鎮", products: [
          { series: "Lote", yearTarget: 349692, monthly: [76000, 0, 0, 76000, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 43396, monthly: [18857, 18857, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 760, null, null, null, null, null, null, null] }
      ] },
      { name: "快樂藥局(時光身心診所)", channel: "診所(含門前藥局)", system: "", city: "台中市大雅區", products: [
          { series: "Epine", yearTarget: 116130, monthly: [0, 0, 42857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 91956, monthly: [28000, 0, 2619, 30524, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 91356, monthly: [26477, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 90769, monthly: [23271, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 9295, monthly: [0, 0, 2858, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 3714, 743, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1430, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂活藥局(心身美診所)", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Ritalin LA", yearTarget: 249450, monthly: [20477, 4786, 20477, 40954, 4786, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 92771, monthly: [9143, 9143, 9143, 9143, 9143, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14857, 14857, 7429, 11143, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 22508, monthly: [0, 1323, 2648, 2648, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 2095, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "呂健弘診所", channel: "診所(含門前藥局)", system: "", city: "台中市西屯區", products: [
          { series: "Ezole", yearTarget: 123123, monthly: [10666, 10666, 10666, 0, 10380, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 121220, monthly: [10210, 10210, 10210, 0, 10210, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 90250, monthly: [9524, 9524, 9524, 9524, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 12600, monthly: [0, 0, 5238, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "三益藥局(宗益身心科診所)", channel: "診所(含門前藥局)", system: "", city: "台中市烏日區", products: [
          { series: "Ritalin LA", yearTarget: 189516, monthly: [16201, 25521, 30107, 12329, 30107, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 59860, monthly: [4757, 4757, 4757, 4588, 4588, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 47167, monthly: [3733, 3732, 3734, 3248, 3247, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2194, 2194, 2194, 2194, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3575, monthly: [1429, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 1314, monthly: [0, 1315, 0, 0, 1314, null, null, null, null, null, null, null] }
      ] },
      { name: "正心藥局(德仁診所)", channel: "診所(含門前藥局)", system: "", city: "台中市南區", products: [
          { series: "Ezole", yearTarget: 158589, monthly: [41334, 0, 7048, 36952, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 99986, monthly: [0, 8658, 4414, 25800, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 19737, monthly: [2080, 0, 2079, 2027, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5700, monthly: [1905, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 4806, monthly: [961, 0, 961, 1921, 961, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3990, monthly: [1143, 0, 2477, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1486, 0, 0, 1486, 743, null, null, null, null, null, null, null] }
      ] },
      { name: "悅康藥師藥局(康誠診所)", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Epine", yearTarget: 221450, monthly: [18000, 12524, 6285, 25476, 15144, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 25034, monthly: [2666, 3619, 0, 3600, 2666, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 18396, monthly: [0, 0, 0, 0, 3943, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 16674, monthly: [0, 0, 0, 4168, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 7600, 0, 0, 1520, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 0, 4114, 0, null, null, null, null, null, null, null] },
          { series: "Suculin", yearTarget: 3668, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1520, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "本堂診所", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Ritalin LA", yearTarget: 87998, monthly: [31429, 0, 31429, 0, 31429, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 61950, monthly: [15352, 0, 0, 14857, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 47142, monthly: [0, 6228, 0, 4714, 4547, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4457, 5200, 0, 0, 1486, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 7515, monthly: [0, 1667, 0, 0, 1667, null, null, null, null, null, null, null] }
      ] },
      { name: "展新診所", channel: "診所(含門前藥局)", system: "展新", city: "台中市西區", products: [
          { series: "Ritalin LA", yearTarget: 111523, monthly: [0, 23760, 0, 15840, 39600, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 88245, monthly: [0, 10000, 0, 9857, 29571, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 12390, monthly: [0, 0, 0, 3543, 1772, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 7600, 0, 4560, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂得好大藥局(開心房診所)", channel: "診所(含門前藥局)", system: "開心房", city: "台中市西屯區", products: [
          { series: "Ritalin LA", yearTarget: 92464, monthly: [11958, 0, 7972, 13552, 15942, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 56067, monthly: [16952, 0, 0, 7619, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 30780, monthly: [6476, 0, 6476, 0, 2438, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7429, 0, 11143, 3714, 1486, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 18396, monthly: [2629, 0, 2629, 0, 3942, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 8360, monthly: [1524, 0, 1524, 1524, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1781, 0, 1781, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "忘憂森林身心診所", channel: "診所(含門前藥局)", system: "", city: "台中市西屯區", products: [
          { series: "Ezole", yearTarget: 84730, monthly: [3048, 36380, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 74898, monthly: [13143, 13143, 13143, 26286, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 17160, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3714, 3714, 3714, 2229, 1486, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3620, monthly: [0, 0, 0, 3620, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新活力診所", channel: "診所(含門前藥局)", system: "", city: "台中市沙鹿區", products: [
          { series: "Epine", yearTarget: 114810, monthly: [0, 0, 0, 40000, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 33507, monthly: [0, 0, 0, 7657, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 16905, 6248, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 21146, monthly: [0, 0, 4806, 0, 4806, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 3786, 757, null, null, null, null, null, null, null] }
      ] },
      { name: "詹東霖心身診所", channel: "診所(含門前藥局)", system: "", city: "台中市中區", products: [
          { series: "Ezole", yearTarget: 153180, monthly: [16667, 0, 16429, 16429, 16429, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 0, 0, 18095, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 3552, 3552, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 371, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "陽光精神科醫院", channel: "地區醫院", system: "", city: "台中市清水區", products: [
          { series: "Lote", yearTarget: 83349, monthly: [20572, 0, 10286, 15429, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 29590, monthly: [4048, 4048, 0, 4048, 4048, null, null, null, null, null, null, null] },
          { series: "Cospirit", yearTarget: 13338, monthly: [16971, 0, 9257, 7571, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 952, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "紓情診所", channel: "診所(含門前藥局)", system: "", city: "台中市西區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [42905, 42905, 42905, 42905, 42905, null, null, null, null, null, null, null] }
      ] },
      { name: "游文治精神科診所", channel: "診所(含門前藥局)", system: "", city: "台中市西區", products: [
          { series: "Epine", yearTarget: 158167, monthly: [13809, 7524, 11286, 15047, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 2391, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "宏安藥局(和康診所)", channel: "診所(含門前藥局)", system: "", city: "台中市后里區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [12322, 12322, 12322, 24644, 12322, null, null, null, null, null, null, null] },
          { series: "Suculin", yearTarget: 41787, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 21280, monthly: [9333, 0, 9334, 0, 2667, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 13728, monthly: [3048, 0, 0, 2939, 2940, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 8550, monthly: [1715, 0, 0, 0, 1714, null, null, null, null, null, null, null] }
      ] },
      { name: "清濱醫院", channel: "地區醫院", system: "賢德", city: "台中市梧棲區", products: [
          { series: "Fute", yearTarget: 115440, monthly: [22667, 0, 0, 9924, 11334, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 11178, monthly: [4857, 0, 0, 2424, 2424, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3040, 0, 0, 760, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1831, 0, 0, 688, 916, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [1219, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林森醫院", channel: "地區醫院", system: "", city: "台中市西區", products: [
          { series: "Tonex", yearTarget: 129027, monthly: [7269, 14537, 18171, 7269, 7269, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 1429, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "全新生活診所", channel: "診所(含門前藥局)", system: "", city: "台中市西屯區", products: [
          { series: "Epine", yearTarget: 42780, monthly: [6905, 0, 6905, 0, 6905, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 34512, monthly: [0, 4803, 0, 2846, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 20894, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 14454, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 12576, monthly: [0, 2619, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4457, 0, 4457, 1486, 743, null, null, null, null, null, null, null] }
      ] },
      { name: "耀聖藥局(陳耀宗診所)", channel: "診所(含門前藥局)", system: "", city: "南投縣竹山鎮", products: [
          { series: "Exelon", yearTarget: 85772, monthly: [7266, 7267, 4360, 9447, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 39143, monthly: [3086, 3200, 4628, 0, 2971, null, null, null, null, null, null, null] }
      ] },
      { name: "惠元診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "南投縣竹山鎮", products: [
          { series: "Epine", yearTarget: 68550, monthly: [28952, 0, 2666, 14476, 11599, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 22860, monthly: [1905, 0, 3810, 5714, 3810, null, null, null, null, null, null, null] }
      ] },
      { name: "靜和仁愛", channel: "地區醫院", system: "", city: "台中市西區", products: [
          { series: "Ezole", yearTarget: 52266, monthly: [8429, 5619, 1124, 5619, 2248, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [8762, 7010, 5257, 10514, 5257, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 22583, monthly: [29334, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "仁美藥局(欣悅診所台中)", channel: "診所(含門前藥局)", system: "", city: "台中市南區", products: [
          { series: "Ezole", yearTarget: 69552, monthly: [15114, 0, 0, 15114, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 44839, monthly: [0, 8006, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4866, 2432, 2433, 1474, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "蘋果樹藥局(柏羽聯合診所)", channel: "診所(含門前藥局)", system: "", city: "台中市西區", products: [
          { series: "Ritalin LA", yearTarget: 55272, monthly: [4016, 6791, 0, 6720, 5083, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 39032, monthly: [8570, 4285, 0, 4200, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [6080, 0, 3040, 1520, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 0, 2838, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 1140, monthly: [0, 0, 0, 1143, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "致遠身心科診所", channel: "診所(含門前藥局)", system: "", city: "南投縣南投市", products: [
          { series: "Ezole", yearTarget: 58286, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 27930, monthly: [0, 0, 0, 13333, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 0, monthly: [13143, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 1067, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "明燈心身診所", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Ritalin LA", yearTarget: 40421, monthly: [3367, 5052, 0, 7578, 4211, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7543, 3771, 2263, 3017, 1509, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 11400, monthly: [1143, 0, 1143, 1143, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3620, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "劉彥山診所", channel: "診所(含門前藥局)", system: "", city: "台中市西區", products: [
          { series: "Exelon", yearTarget: 35475, monthly: [7200, 0, 11520, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 18171, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "惠良精神科診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "南投縣南投市", products: [
          { series: "Epine", yearTarget: 27360, monthly: [7619, 0, 0, 7619, 15238, null, null, null, null, null, null, null] }
      ] },
      { name: "昕晴診所", channel: "診所(含門前藥局)", system: "", city: "台中市西屯區", products: [
          { series: "Lote", yearTarget: 30222, monthly: [26286, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "得安藥局(后里診所)", channel: "診所(含門前藥局)", system: "", city: "台中市后里區", products: [
          { series: "Exelon", yearTarget: 39391, monthly: [6516, 1447, 723, 5069, 2895, null, null, null, null, null, null, null] }
      ] },
      { name: "懷信中西藥局(後龍診所)", channel: "診所(含門前藥局)", system: "", city: "苗栗縣後龍鎮", products: [
          { series: "Epine", yearTarget: 39071, monthly: [0, 4620, 0, 7000, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 2857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 1714, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "喜樂藥局(惠承診所)", channel: "診所(含門前藥局)", system: "賈惠洲", city: "南投縣埔里鎮", products: [
          { series: "Epine", yearTarget: 33375, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 12364, monthly: [0, 0, 0, 0, 3371, null, null, null, null, null, null, null] }
      ] },
      { name: "永佳診所", channel: "診所(含門前藥局)", system: "", city: "苗栗縣竹南鎮", products: [
          { series: "Lote", yearTarget: 32637, monthly: [0, 14191, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "文聖診所", channel: "診所(含門前藥局)", system: "", city: "台中市南區", products: [
          { series: "Fute", yearTarget: 24435, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 14880, monthly: [0, 3714, 0, 3714, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "立好藥局(晨安診所)", channel: "診所(含門前藥局)", system: "", city: "台中市沙鹿區", products: [
          { series: "Ritalin LA", yearTarget: 31730, monthly: [8137, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1486, 0, 0, 743, null, null, null, null, null, null, null] }
      ] },
      { name: "沙鹿合和健保藥局(醫達人診所)", channel: "診所(含門前藥局)", system: "", city: "台中市沙鹿區", products: [
          { series: "Ritalin LA", yearTarget: 27489, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 3285, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 1140, 0, 1140, null, null, null, null, null, null, null] }
      ] },
      { name: "李潤宇診所", channel: "診所(含門前藥局)", system: "", city: "台中市西區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [9714, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 6820, monthly: [8666, 0, 0, 2477, 0, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 0, 4525, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "施以中診所", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Fute", yearTarget: 21420, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 10725, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "謝明哲診所", channel: "診所(含門前藥局)", system: "", city: "南投縣竹山鎮", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [6571, 3286, 0, 6571, 0, null, null, null, null, null, null, null] },
          { series: "Suculin", yearTarget: 10746, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1791, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "梧棲明德", channel: "地區醫院", system: "賢德", city: "台中市梧棲區", products: [
          { series: "Fute", yearTarget: 21250, monthly: [0, 0, 0, 4962, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 0, 1219, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 688, 458, null, null, null, null, null, null, null] }
      ] },
      { name: "澄清復健", channel: "地區醫院", system: "", city: "台中市西屯區", products: [
          { series: "Tonex", yearTarget: 14648, monthly: [0, 3662, 0, 0, 3662, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 3620, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 357, 0, 0, null, null, null, null, null, null, null] },
          { series: "Pane", yearTarget: 336, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "瑞東診所", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Exelon Patch", yearTarget: 24456, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "烏日澄清", channel: "地區醫院", system: "", city: "台中市烏日區", products: [
          { series: "Tonex", yearTarget: 14757, monthly: [0, 3690, 0, 0, 3690, null, null, null, null, null, null, null] }
      ] },
      { name: "仁厚藥局(松群診所)", channel: "診所(含門前藥局)", system: "", city: "台中市烏日區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 0, 21257, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "康安新美藥局(益康診所)", channel: "診所(含門前藥局)", system: "", city: "苗栗縣苗栗市", products: [
          { series: "Lote", yearTarget: 10668, monthly: [0, 9143, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "得心身醫學向上診所", channel: "診所(含門前藥局)", system: "得立", city: "台中市南屯區", products: [
          { series: "Ritalin LA", yearTarget: 4628, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 2420, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1995, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1430, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 1319, monthly: [2640, 0, 0, 0, 2613, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 457, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "清泉醫院", channel: "地區醫院", system: "", city: "台中市大雅區", products: [
          { series: "Tone", yearTarget: 6175, monthly: [1904, 0, 0, 1905, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5490, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "清水回生診所", channel: "診所(含門前藥局)", system: "", city: "台中市清水區", products: [
          { series: "Bisadyl Supp", yearTarget: 10990, monthly: [0, 0, 0, 3143, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林令世診所", channel: "診所(含門前藥局)", system: "", city: "台中市沙鹿區", products: [
          { series: "Tone", yearTarget: 6175, monthly: [5715, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "佛教正德醫院", channel: "地區醫院", system: "中榮", city: "台中市南屯區", products: [
          { series: "Exelon", yearTarget: 7226, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 2286, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 0, 2096, null, null, null, null, null, null, null] }
      ] },
      { name: "林佳輝診所", channel: "診所(含門前藥局)", system: "", city: "台中市清水區", products: [
          { series: "Fute", yearTarget: 4650, monthly: [3715, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "草鞋墩神經診所", channel: "診所(含門前藥局)", system: "", city: "南投縣草屯鎮", products: [
          { series: "Tonex", yearTarget: 0, monthly: [4114, 0, 0, 4114, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "全佑藥局(全佑家庭醫學科診所)", channel: "診所(含門前藥局)", system: "", city: "台中市清水區", products: [
          { series: "Fute", yearTarget: 7695, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "圓情居身心靈診所", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [2971, 2971, 743, 743, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "保仁藥局(蕭天讚小兒科診所)", channel: "診所(含門前藥局)", system: "", city: "台中市后里區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 4896, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 2240, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "惠昕身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "台中市西區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [3729, 3271, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "天恩藥局(天恩診所)", channel: "診所(含門前藥局)", system: "", city: "苗栗縣苗栗市", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [9857, -6571, 0, 0, 3286, null, null, null, null, null, null, null] }
      ] },
      { name: "慧寧藥局(家珊診所)", channel: "診所(含門前藥局)", system: "", city: "台中市南區", products: [
          { series: "Exprexa", yearTarget: 2134, monthly: [0, 2133, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 369, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "梁力仁診所", channel: "診所(含門前藥局)", system: "", city: "南投縣南投市", products: [
          { series: "Tone", yearTarget: 4445, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "德恩藥師藥局(陳家驊骨科診所)", channel: "診所(含門前藥局)", system: "", city: "台中市西區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 0, 0, 4114, null, null, null, null, null, null, null] }
      ] },
      { name: "東亞永福藥局(建安復健科診所)", channel: "診所(含門前藥局)", system: "", city: "台中市西屯區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 0, 4114, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "王壽山骨科診所", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 0, 4114, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "祥和大藥局(禾欣骨科診所)", channel: "診所(含門前藥局)", system: "", city: "台中市西屯區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [0, 0, 0, 0, 4114, null, null, null, null, null, null, null] }
      ] },
      { name: "第一醫院", channel: "地區醫院", system: "保健安", city: "台中市中區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 1643, 1971, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 228, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "中美偕診所", channel: "診所(含門前藥局)", system: "", city: "台中市西區", products: [
          { series: "Epine", yearTarget: 3800, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "正開心大藥局(挺開朗身心診所)", channel: "診所(含門前藥局)", system: "開心房", city: "台中市霧峰區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 3384, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 371, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "昱宏診所", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [3581, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "水里社區基督聯合診所", channel: "診所(含門前藥局)", system: "埔基", city: "南投縣水里鄉", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [569, 0, 284, 569, 855, null, null, null, null, null, null, null] }
      ] },
      { name: "吳東洀診所", channel: "診所(含門前藥局)", system: "", city: "台中市西區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 760, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "莫亞皮膚科診所", channel: "診所(含門前藥局)", system: "", city: "台中市西屯區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 228, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "賈光耀", group: "Div3", customers: [
      { name: "振心身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Ritalin LA", yearTarget: 339641, monthly: [58629, 16596, 37691, 35364, 38994, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 91405, monthly: [19685, 13124, 0, 15114, 15114, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 87266, monthly: [8840, 6800, 0, 0, 3400, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 71936, monthly: [11238, 11238, 4495, 8991, 11238, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 44640, monthly: [9905, 0, 4953, 4952, 4952, null, null, null, null, null, null, null] }
      ] },
      { name: "心晴診所桃園", channel: "診所(含門前藥局)", system: "心寧", city: "桃園市桃園區", products: [
          { series: "Ritalin LA", yearTarget: 627718, monthly: [0, 0, 0, 82985, 52046, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 47914, monthly: [12438, 0, 0, 0, 5498, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 45288, monthly: [0, 0, 10065, 10065, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14571, 0, 14571, 1457, 4371, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 8540, monthly: [2438, 0, 0, 0, 3657, null, null, null, null, null, null, null] }
      ] },
      { name: "心寧診所", channel: "診所(含門前藥局)", system: "心寧", city: "桃園市桃園區", products: [
          { series: "Ritalin LA", yearTarget: 498035, monthly: [122086, 0, 42972, 79114, 42972, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 57344, monthly: [18939, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 47804, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2914, 0, 0, 1457, 729, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 3548, monthly: [3040, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安和藥局(交感身心診所)", channel: "診所(含門前藥局)", system: "1B", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 331617, monthly: [6818, 65852, 116097, 0, 166343, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14742, 7371, 7415, 1865, 1492, null, null, null, null, null, null, null] }
      ] },
      { name: "崇光身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Ezole", yearTarget: 114594, monthly: [12571, 6286, 6286, 12572, 12571, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 95503, monthly: [8057, 8783, 11721, 16112, 20504, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 88094, monthly: [11833, 5571, 8119, 16238, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 58949, monthly: [0, 0, 15514, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 48421, monthly: [11810, 0, 11810, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 33124, monthly: [0, 0, 3293, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 15498, monthly: [0, 0, 8027, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2880, 2880, 2880, 2880, 720, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5985, monthly: [0, 0, 0, 0, 1333, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1543, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "靜和藥局(育禾身心診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "桃園市蘆竹區", products: [
          { series: "Ritalin LA", yearTarget: 356106, monthly: [38472, 30777, 52488, 31511, 37099, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 32000, monthly: [4000, 4000, 4000, 0, 6000, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7514, 15029, 3757, 3757, 1503, null, null, null, null, null, null, null] }
      ] },
      { name: "翰生藥局(翰生診所)", channel: "診所(含門前藥局)", system: "", city: "新北市板橋區", products: [
          { series: "Ritalin LA", yearTarget: 425451, monthly: [0, 0, 0, 880, 90902, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 16998, monthly: [0, 0, 0, 0, 1286, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 12138, monthly: [0, 0, 0, 0, 3714, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 6034, null, null, null, null, null, null, null] }
      ] },
      { name: "夏凱納診所", channel: "診所(含門前藥局)", system: "", city: "台北市大同區", products: [
          { series: "Ritalin LA", yearTarget: 284349, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 149512, monthly: [0, 0, 0, 0, 25620, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 31512, monthly: [0, 0, 0, 2714, 1809, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 15233, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 9181, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 1440, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "悅藥坊人文藥局(家恩診所)", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Epine", yearTarget: 339990, monthly: [0, 59715, 38572, 0, 27933, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14743, 0, 0, 5160, 2211, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 10191, monthly: [0, 0, 5489, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "欣悅診所桃園", channel: "診所(含門前藥局)", system: "心寧", city: "桃園市蘆竹區", products: [
          { series: "Ritalin LA", yearTarget: 259808, monthly: [0, 0, 52046, 0, 6446, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 94350, monthly: [12581, 0, 0, 0, 12581, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 14571, 7286, 2914, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 17524, monthly: [2543, 7858, 0, 2772, 2772, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 3548, monthly: [0, 1013, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "正文身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Ritalin LA", yearTarget: 96882, monthly: [14462, 0, 28926, 0, 21694, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 63535, monthly: [13143, 0, 26285, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 61740, monthly: [0, 9524, 8953, 0, 9228, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 40635, monthly: [0, 7524, 0, 0, 9000, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [13143, 0, 13143, 5257, 3285, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 9257, 0, 0, 15429, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 4952, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 1319, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "一德身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Ezole", yearTarget: 88770, monthly: [0, 0, 21429, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 78638, monthly: [0, 25238, 0, 0, 28106, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 59472, monthly: [0, 12390, 0, 18586, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 33480, monthly: [0, 0, 12381, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 21960, monthly: [0, 0, 4392, 0, 4392, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 15200, 3800, 1520, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 16553, monthly: [0, 0, 7330, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "仁濟新莊", channel: "地區醫院", system: "仁濟", city: "新北市新莊區", products: [
          { series: "Exprexa", yearTarget: 181771, monthly: [0, 0, 0, 18263, 11606, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 82968, monthly: [0, 0, 0, 8679, 6000, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 71873, monthly: [0, 0, 0, 5819, 6983, null, null, null, null, null, null, null] },
          { series: "Cospirit", yearTarget: 0, monthly: [0, 0, 0, 9024, 9238, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 3185, 3475, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 2195, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新華興藥局(林瑞家診所)", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Exelon", yearTarget: 273354, monthly: [9969, 26346, 17091, 14953, 5696, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 12495, 0, 12495, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安興精神科診所", channel: "診所(含門前藥局)", system: "安興", city: "新北市三重區", products: [
          { series: "Ritalin LA", yearTarget: 129715, monthly: [0, 0, 0, 28368, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 60667, monthly: [0, 0, 0, 0, 7429, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 43312, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 40954, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 21441, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 8364, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "晴美身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市新莊區", products: [
          { series: "Ritalin LA", yearTarget: 314463, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 7309, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 3557, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 933, monthly: [0, 0, 0, 0, 2667, null, null, null, null, null, null, null] }
      ] },
      { name: "大一藥局(聖家小兒科診所)", channel: "診所(含門前藥局)", system: "1B", city: "桃園市桃園區", products: [
          { series: "Exelon", yearTarget: 171807, monthly: [13606, 13606, 29185, 0, 29185, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [10812, 2012, 0, 9614, 2057, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [4571, 4572, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 6824, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 0, monthly: [0, 0, 0, 5765, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 4100, monthly: [0, 0, 0, 8296, 0, null, null, null, null, null, null, null] },
          { series: "Bisadyl", yearTarget: 0, monthly: [0, 0, 3620, 0, 0, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 0, monthly: [0, 0, 0, 0, 1867, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 0, monthly: [0, 0, 904, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "和沛身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "台北市文山區", products: [
          { series: "Ritalin LA", yearTarget: 145247, monthly: [0, 0, 0, 48023, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 66827, monthly: [0, 0, 0, 11619, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 31265, monthly: [0, 0, 0, 9687, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 3571, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "佳鄰藥局(家齡診所)", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 94870, monthly: [15428, 0, 15428, 0, 15428, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 63058, monthly: [4388, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 60529, monthly: [4731, 0, 0, 4490, 7483, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2972, 0, 2972, 1486, 743, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3428, 0, 1715, 1715, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1971, 0, 986, 1971, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1029, 2743, 0, 1029, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大安身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 117889, monthly: [7557, 15114, 0, 15114, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 52360, monthly: [0, 0, 14286, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 22286, 14857, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 37072, monthly: [0, 0, 0, 11914, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "心安藥局(育心身心精神科診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "新北市新莊區", products: [
          { series: "Ritalin LA", yearTarget: 140326, monthly: [0, 0, 0, 0, 14017, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 106301, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 27711, monthly: [0, 0, 0, 0, 8000, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 751, null, null, null, null, null, null, null] }
      ] },
      { name: "欣悅藥局(美麗心成人兒童精神科診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "新北市三重區", products: [
          { series: "Ritalin LA", yearTarget: 180378, monthly: [0, 0, 0, 66503, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 48493, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "喜樂藥局(松德精神科診所)", channel: "診所(含門前藥局)", system: "", city: "台北市信義區", products: [
          { series: "Ritalin LA", yearTarget: 82853, monthly: [13606, 0, 8212, 13469, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [24342, 0, 12171, 7303, 4869, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 32000, monthly: [4000, 0, 4000, 4000, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 27861, monthly: [3151, 0, 6303, 0, 2994, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 22097, monthly: [0, 2870, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "木柵身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市文山區", products: [
          { series: "Ritalin LA", yearTarget: 120670, monthly: [0, 0, 0, 122657, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 24895, monthly: [0, 0, 0, 3429, 3429, null, null, null, null, null, null, null] }
      ] },
      { name: "祐康藥局(養心診所)", channel: "診所(含門前藥局)", system: "", city: "新北市鶯歌區", products: [
          { series: "Lote", yearTarget: 197090, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 47081, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安泰藥局(安心診所)", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Exprexa", yearTarget: 89321, monthly: [0, 0, 0, 0, 6784, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 43094, monthly: [0, 0, 22286, 0, 25966, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 10428, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2280, 0, 0, 1520, 1520, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 4967, monthly: [9028, 0, 0, 6772, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂心藥局(順心診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "新北市板橋區", products: [
          { series: "Ritalin LA", yearTarget: 100712, monthly: [0, 0, 0, 15388, 29405, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 41961, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 21084, monthly: [0, 0, 0, 4000, 2400, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 3757, 751, null, null, null, null, null, null, null] }
      ] },
      { name: "心世界身心精神科診所", channel: "診所(含門前藥局)", system: "", city: "新北市板橋區", products: [
          { series: "Lote", yearTarget: 105394, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 43540, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 32153, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 16551, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 13089, monthly: [0, 0, 0, 0, 3993, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 2892, null, null, null, null, null, null, null] }
      ] },
      { name: "龍霖身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市林口區", products: [
          { series: "Ritalin LA", yearTarget: 163572, monthly: [0, 0, 0, 22637, 22637, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 2229, 2229, null, null, null, null, null, null, null] }
      ] },
      { name: "福田診所", channel: "診所(含門前藥局)", system: "", city: "新北市蘆洲區", products: [
          { series: "Ritalin LA", yearTarget: 166694, monthly: [0, 0, 0, 0, 38443, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "伯特利身心診所", channel: "診所(含門前藥局)", system: "伯特利", city: "台北市信義區", products: [
          { series: "Ritalin LA", yearTarget: 65896, monthly: [7403, 7403, 0, 7403, 11105, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 49530, monthly: [0, 0, 0, 11429, 11429, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14857, 14857, 0, 2229, 2972, null, null, null, null, null, null, null] }
      ] },
      { name: "藥濟藥局(欣慈診所)", channel: "診所(含門前藥局)", system: "崇文", city: "新北市鶯歌區", products: [
          { series: "Ritalin LA", yearTarget: 61988, monthly: [0, 0, 0, 5887, 6076, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 36747, monthly: [0, 0, 0, 5905, 0, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 27672, monthly: [0, 0, 0, 4190, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 10571, 10571, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 16979, monthly: [0, 0, 0, 0, 2191, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 0, 0, 8095, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1808, 723, null, null, null, null, null, null, null] }
      ] },
      { name: "昀樂身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市新莊區", products: [
          { series: "Ritalin LA", yearTarget: 57225, monthly: [0, 0, 0, 0, 75549, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 9865, monthly: [0, 0, 0, 0, 15714, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 7428, 2229, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 5420, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5342, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 2797, monthly: [0, 0, 0, 0, 12381, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 2691, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1039, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "福全身心科診所", channel: "診所(含門前藥局)", system: "", city: "台北市萬華區", products: [
          { series: "Epine", yearTarget: 55564, monthly: [0, 0, 0, 0, 30952, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 55366, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 28572, 10000, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 9729, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "伯特利敦南身心診所", channel: "診所(含門前藥局)", system: "伯特利", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 99214, monthly: [0, 14805, 14805, 0, 22208, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 18288, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2971, 2971, 2971, 5942, 1486, null, null, null, null, null, null, null] }
      ] },
      { name: "西園醫院", channel: "地區醫院", system: "", city: "台北市萬華區", products: [
          { series: "Lote", yearTarget: 141210, monthly: [0, 0, 0, 18324, 14095, null, null, null, null, null, null, null] }
      ] },
      { name: "欣泉身心診所", channel: "診所(含門前藥局)", system: "張邦彥", city: "新北市林口區", products: [
          { series: "Lote", yearTarget: 78292, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 58369, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 29658, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "吳俊毅身心精神科診所", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Ezole", yearTarget: 78604, monthly: [0, 0, 0, 34286, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 27370, monthly: [0, 0, 0, 23809, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "恆友精神科診所", channel: "診所(含門前藥局)", system: "安興", city: "新北市新莊區", products: [
          { series: "Ritalin LA", yearTarget: 79591, monthly: [0, 0, 0, 15702, 12666, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 13984, monthly: [0, 0, 0, 3714, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 11727, monthly: [0, 0, 0, 0, 3181, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 10121, monthly: [0, 0, 0, 1619, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 6714, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 2390, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "劉建宏診所", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 95054, monthly: [20366, 0, 12729, 0, 10183, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 2629, 3614, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [746, 0, 373, 0, 746, null, null, null, null, null, null, null] }
      ] },
      { name: "吾家藥局(振芝映煦心身醫學診所)", channel: "診所(含門前藥局)", system: "振芝", city: "台北市大安區", products: [
          { series: "Ezole", yearTarget: 38758, monthly: [9466, 0, 7867, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 37584, monthly: [7829, 7829, 4697, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 6190, monthly: [2476, 0, 1239, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 2860, monthly: [0, 0, 1429, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2228, 0, 520, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1330, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "雷亞診所", channel: "診所(含門前藥局)", system: "", city: "桃園市八德區", products: [
          { series: "Exprexa", yearTarget: 39200, monthly: [0, 4480, 0, 6704, 2234, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 26418, monthly: [0, 5033, 2516, 2516, 1257, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 14654, monthly: [0, 0, 0, 0, 4305, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 12989, monthly: [0, 2095, 2095, 0, 2091, null, null, null, null, null, null, null] }
      ] },
      { name: "新店開心診所", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Ritalin LA", yearTarget: 124093, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "治聿身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Ritalin LA", yearTarget: 51716, monthly: [0, 0, 0, 15428, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 24687, monthly: [0, 0, 0, 9107, 4555, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 11896, monthly: [0, 0, 0, 0, 3587, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 743, 743, null, null, null, null, null, null, null] }
      ] },
      { name: "芯安藥局(欣寧身心診所)", channel: "診所(含門前藥局)", system: "", city: "新北市三峽區", products: [
          { series: "Ritalin LA", yearTarget: 121244, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "悅情身心科診所", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Ritalin LA", yearTarget: 75136, monthly: [6956, 0, 0, 6956, 9274, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [6228, 0, 0, 6228, 1868, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3714, 0, 1857, 0, 743, null, null, null, null, null, null, null] }
      ] },
      { name: "看見佳醫診所", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 70250, monthly: [17492, -17492, 0, 18720, 12896, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [36857, -22114, 0, 0, 2237, null, null, null, null, null, null, null] }
      ] },
      { name: "合誼身心診所", channel: "診所(含門前藥局)", system: "進安", city: "新北市蘆洲區", products: [
          { series: "Ezole", yearTarget: 45318, monthly: [0, 0, 0, 0, 7429, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 10990, monthly: [0, 0, 0, 1391, 1391, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 6428, 714, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 7024, monthly: [0, 0, 0, 15715, 11606, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 5324, monthly: [0, 0, 0, 2829, 1414, null, null, null, null, null, null, null] }
      ] },
      { name: "均樂身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Ritalin LA", yearTarget: 65558, monthly: [0, 0, 0, 31429, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 9285, 5000, null, null, null, null, null, null, null] }
      ] },
      { name: "長源身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市信義區", products: [
          { series: "Ezole", yearTarget: 48438, monthly: [0, 0, 8191, 0, 16666, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 8100, monthly: [5048, 0, 5238, 0, 11905, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2971, 0, 0, 1486, 743, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 1238, monthly: [0, 0, 1238, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "公館得心身醫學診所", channel: "診所(含門前藥局)", system: "得立", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 24542, monthly: [3085, 0, 0, 6053, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 21760, monthly: [2639, 0, 2640, 0, 2640, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 14970, monthly: [1500, 3991, 0, 1500, 3991, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 8603, monthly: [0, 2457, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1440, 1440, 2160, 720, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 914, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "放開心身心診所", channel: "診所(含門前藥局)", system: "好心情", city: "新北市板橋區", products: [
          { series: "Fute", yearTarget: 59231, monthly: [0, 0, 0, 0, 27428, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 7831, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 5922, monthly: [0, 0, 0, 0, 2858, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 1723, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 1486, null, null, null, null, null, null, null] }
      ] },
      { name: "佳醫-新泰醫院", channel: "地區醫院", system: "佳醫集康", city: "新北市新莊區", products: [
          { series: "Epram", yearTarget: 64938, monthly: [0, 0, 0, 5286, 7048, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 19227, monthly: [0, 0, 0, 3648, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 3157, 3158, null, null, null, null, null, null, null] }
      ] },
      { name: "新店北新身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Ritalin LA", yearTarget: 45134, monthly: [0, 0, 0, 6989, 3106, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 37885, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 5714, 2858, null, null, null, null, null, null, null] }
      ] },
      { name: "萬華維沐身心診所", channel: "診所(含門前藥局)", system: "張維仁", city: "台北市萬華區", products: [
          { series: "Ritalin LA", yearTarget: 59949, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 26658, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 5342, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 714, null, null, null, null, null, null, null] }
      ] },
      { name: "心翔身心精神科診所", channel: "診所(含門前藥局)", system: "林威廷", city: "新北市新莊區", products: [
          { series: "Lote", yearTarget: 37018, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 36483, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 11148, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 6019, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "柏康藥局(丁良文內科診所)", channel: "診所(含門前藥局)", system: "", city: "新北市蘆洲區", products: [
          { series: "Fute", yearTarget: 32733, monthly: [0, 0, 0, 6286, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 16429, 2590, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 16603, monthly: [0, 0, 0, 0, 5238, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 6070, monthly: [0, 0, 0, 2477, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "大心診所百年館", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 29452, monthly: [4117, 4116, 12433, 12476, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [6634, 7372, 7372, 1865, 746, null, null, null, null, null, null, null] }
      ] },
      { name: "宏恩藥局(怡心診所)", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [8232, 2727, 12835, 15453, 10182, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 0, monthly: [5653, 3533, 4946, 4239, 2826, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3686, 2949, 1474, 1474, 1474, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 0, 1702, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [625, 0, 0, 625, 0, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 0, monthly: [0, 0, 0, 904, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 0, 0, 892, null, null, null, null, null, null, null] }
      ] },
      { name: "捷思診所", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Lote", yearTarget: 21420, monthly: [8333, 0, 4762, 7143, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 18707, monthly: [3514, 4019, 4685, 1952, 3209, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 1280, monthly: [1280, 0, 0, 2560, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "書田診所", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Slatone", yearTarget: 29750, monthly: [0, 5257, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [14476, 0, 0, 0, 3618, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [3571, 0, 0, 2500, 2411, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5670, monthly: [4857, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "誠心身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "台北市文山區", products: [
          { series: "Ritalin LA", yearTarget: 32404, monthly: [0, 0, 0, 23297, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 7428, 7429, null, null, null, null, null, null, null] }
      ] },
      { name: "楊思亮診所", channel: "診所(含門前藥局)", system: "楊思亮", city: "新北市板橋區", products: [
          { series: "Ezole", yearTarget: 38221, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 16551, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 10357, monthly: [0, 0, 0, 3943, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 933, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "蘇宗偉身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市龜山區", products: [
          { series: "Epine", yearTarget: 30106, monthly: [0, 0, 0, 9524, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 23065, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 5286, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新信義藥局(欣彥身心診所)", channel: "診所(含門前藥局)", system: "1B", city: "新北市蘆洲區", products: [
          { series: "Epine", yearTarget: 24438, monthly: [0, 0, 0, 0, 2286, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 15361, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 7063, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 5150, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 4698, monthly: [0, 0, 0, 0, 1247, null, null, null, null, null, null, null] }
      ] },
      { name: "楊孟達身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "新北市板橋區", products: [
          { series: "Ritalin LA", yearTarget: 30802, monthly: [0, 0, 0, 0, 11075, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 8043, monthly: [0, 0, 0, 0, 5295, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 2800, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 2145, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "喜樂藥局(喜悅診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "新北市樹林區", products: [
          { series: "Ezole", yearTarget: 28915, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 11604, monthly: [0, 0, 0, 8831, 6743, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1503, 2254, null, null, null, null, null, null, null] }
      ] },
      { name: "向陽澄思診所", channel: "診所(含門前藥局)", system: "張邦彥", city: "新北市板橋區", products: [
          { series: "Ezole", yearTarget: 29027, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 25003, monthly: [0, 0, 0, 0, 2685, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "和信藥師藥局(聖洸身心診所)", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Fute", yearTarget: 36575, monthly: [6667, 0, 0, 0, 6667, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 5005, monthly: [0, 0, 0, 0, 2856, null, null, null, null, null, null, null] }
      ] },
      { name: "向陽身心診所", channel: "診所(含門前藥局)", system: "張邦彥", city: "台北市大安區", products: [
          { series: "Ezole", yearTarget: 19693, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 15996, monthly: [4000, 0, 4000, 2667, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2429, 0, 2429, 1071, 1429, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1789, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "沐光藥局(沐陽身心科診所)", channel: "診所(含門前藥局)", system: "", city: "新北市永和區", products: [
          { series: "Ezole", yearTarget: 0, monthly: [11428, 0, 9286, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 0, monthly: [3866, 4800, 6266, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [3020, 5657, 1557, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1440, 720, 1872, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [1335, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [1190, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "太陽雨信義身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市信義區", products: [
          { series: "Ritalin LA", yearTarget: 18207, monthly: [2439, 0, 3893, 3088, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2600, 1857, 2229, 3715, 743, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [1905, 0, 0, 0, 1900, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 3426, monthly: [1714, 0, 0, 2853, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "好心情身心診所", channel: "診所(含門前藥局)", system: "好心情", city: "新北市新莊區", products: [
          { series: "Fute", yearTarget: 32619, monthly: [0, 0, 0, 9144, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 4518, monthly: [0, 0, 0, 1199, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 1486, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 790, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "陳信任精神科診所", channel: "診所(含門前藥局)", system: "", city: "新北市三重區", products: [
          { series: "Epine", yearTarget: 35376, monthly: [0, 0, 0, 8096, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 3396, monthly: [0, 0, 0, 2279, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂開懷健保藥局(開心生活診所)", channel: "診所(含門前藥局)", system: "開心房", city: "新北市蘆洲區", products: [
          { series: "Ezole", yearTarget: 6599, monthly: [0, 0, 0, 7620, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 6532, monthly: [0, 0, 0, 23914, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 3714, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "上安藥局(晨新診所)", channel: "診所(含門前藥局)", system: "呂慶城", city: "桃園市蘆竹區", products: [
          { series: "Ritalin LA", yearTarget: 28130, monthly: [0, 0, 6430, 2412, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 0, 3286, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 897, 0, 748, null, null, null, null, null, null, null] }
      ] },
      { name: "晴天身心診所", channel: "診所(含門前藥局)", system: "張邦彥", city: "台北市大安區", products: [
          { series: "Lote", yearTarget: 23994, monthly: [5333, 0, 0, 5333, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2857, 0, 0, 1429, 1429, null, null, null, null, null, null, null] }
      ] },
      { name: "圓舞曲藥局(圓舞曲身心診所)", channel: "診所(含門前藥局)", system: "", city: "台北市信義區", products: [
          { series: "Exelon Patch", yearTarget: 0, monthly: [8485, 0, 7200, 0, 2846, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [2857, 0, 0, 8572, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 0, monthly: [0, 0, 1381, 1381, 4143, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 0, monthly: [2073, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [1334, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "奇點藥局(芯悅診所)", channel: "診所(含門前藥局)", system: "", city: "桃園市蘆竹區", products: [
          { series: "Tone", yearTarget: 25840, monthly: [4571, 0, 0, 0, 9142, null, null, null, null, null, null, null] }
      ] },
      { name: "主愛心靈診所", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 3571, 3124, 3571, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5200, 0, 1486, 1486, 1486, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 3864, monthly: [0, 0, 3094, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 790, 790, 1581, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 2740, monthly: [0, 1372, 2743, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 1714, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "當陽之境身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Ritalin LA", yearTarget: 29469, monthly: [2327, 0, 0, 2327, 2327, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1429, 0, 0, 71, 357, null, null, null, null, null, null, null] }
      ] },
      { name: "擁抱心身醫學診所", channel: "診所(含門前藥局)", system: "抱抱", city: "桃園市桃園區", products: [
          { series: "Ezole", yearTarget: 17332, monthly: [0, 0, 2476, 0, 7429, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 3150, monthly: [2095, 0, 0, 0, 2095, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [714, 0, 714, 714, 714, null, null, null, null, null, null, null] }
      ] },
      { name: "林口抱抱心身醫學診所", channel: "診所(含門前藥局)", system: "抱抱", city: "新北市林口區", products: [
          { series: "Ritalin LA", yearTarget: 7153, monthly: [0, 0, 0, 10586, 9171, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 3262, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 2372, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 714, 1071, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 0, 313, null, null, null, null, null, null, null] }
      ] },
      { name: "詠美身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Epine", yearTarget: 20580, monthly: [5143, 0, 0, 0, 1710, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 0, monthly: [0, 0, 0, 0, 5714, null, null, null, null, null, null, null] }
      ] },
      { name: "仁濟台北", channel: "地區醫院", system: "仁濟", city: "台北市萬華區", products: [
          { series: "Exelon Patch", yearTarget: 18643, monthly: [0, 0, 0, 10382, 3797, null, null, null, null, null, null, null] }
      ] },
      { name: "溫時光診所", channel: "診所(含門前藥局)", system: "", city: "新北市新莊區", products: [
          { series: "Ezole", yearTarget: 26837, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 2229, null, null, null, null, null, null, null] }
      ] },
      { name: "宏通診所", channel: "診所(含門前藥局)", system: "", city: "新北市三重區", products: [
          { series: "Epine", yearTarget: 15868, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 6571, 6571, null, null, null, null, null, null, null] }
      ] },
      { name: "王湘琦身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市三峽區", products: [
          { series: "Fute", yearTarget: 10165, monthly: [0, 0, 0, 0, 4057, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 8746, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 4202, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "萬華醫院", channel: "地區醫院", system: "", city: "台北市萬華區", products: [
          { series: "Exelon Patch", yearTarget: 25179, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "中永和身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "新北市永和區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [14000, 10500, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林南診所", channel: "診所(含門前藥局)", system: "", city: "新北市三重區", products: [
          { series: "Ritalin LA", yearTarget: 10990, monthly: [0, 0, 0, 1342, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 10241, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 357, 357, null, null, null, null, null, null, null] }
      ] },
      { name: "新田身心診所龜山", channel: "診所(含門前藥局)", system: "新田", city: "桃園市龜山區", products: [
          { series: "Ritalin LA", yearTarget: 18962, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "居善醫院", channel: "地區醫院", system: "", city: "桃園市大園區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [9257, 0, 0, 0, 9257, null, null, null, null, null, null, null] }
      ] },
      { name: "興隆藥局(方舟身心診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "台北市文山區", products: [
          { series: "Ritalin LA", yearTarget: 15274, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 751, 1503, null, null, null, null, null, null, null] }
      ] },
      { name: "壬康診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "新北市蘆洲區", products: [
          { series: "Lote", yearTarget: 5078, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 4951, monthly: [0, 0, 0, 0, 739, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 4016, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 700, null, null, null, null, null, null, null] }
      ] },
      { name: "中山抱抱心身醫學診所", channel: "診所(含門前藥局)", system: "抱抱", city: "台北市大同區", products: [
          { series: "Fute", yearTarget: 8012, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 4000, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 1864, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 357, 357, null, null, null, null, null, null, null] }
      ] },
      { name: "明心診所", channel: "診所(含門前藥局)", system: "", city: "新北市永和區", products: [
          { series: "Fute", yearTarget: 4599, monthly: [2666, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 4541, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [2562, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "後埔永江診所", channel: "診所(含門前藥局)", system: "", city: "新北市板橋區", products: [
          { series: "Fute", yearTarget: 10119, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 991, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林青穀診所", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [5029, 0, 0, 3352, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "沐林親子診所", channel: "診所(含門前藥局)", system: "", city: "新北市板橋區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 0, 7977, null, null, null, null, null, null, null] }
      ] },
      { name: "幸福身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "新北市板橋區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1400, 1400, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 2145, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 1005, monthly: [0, 0, 0, 1324, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "龍城藥局(周孫元診所)", channel: "診所(含門前藥局)", system: "", city: "桃園市桃園區", products: [
          { series: "Fute", yearTarget: 5130, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "和安藥局(晨益診所)", channel: "診所(含門前藥局)", system: "呂慶城", city: "新北市板橋區", products: [
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 0, 0, 2476, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 1607, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 373, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安法診所", channel: "診所(含門前藥局)", system: "", city: "台北市大安區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [2172, 0, 0, 2172, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "富康藥局(建順診所)", channel: "診所(含門前藥局)", system: "", city: "台北市大同區", products: [
          { series: "Lote", yearTarget: 3658, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "萬華身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "台北市萬華區", products: [
          { series: "Epine", yearTarget: 1005, monthly: [0, 0, 0, 2647, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "惠生診所", channel: "診所(含門前藥局)", system: "", city: "新北市永和區", products: [
          { series: "Epine", yearTarget: 3246, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 375, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "正欣藥局(正興診所)", channel: "診所(含門前藥局)", system: "", city: "新北市三重區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 3286, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "八德身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "桃園市八德區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 2800, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "永安藥局桃園(晨暘診所)", channel: "診所(含門前藥局)", system: "呂慶城", city: "桃園市桃園區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 0, 1607, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 822, null, null, null, null, null, null, null] }
      ] },
      { name: "里安診所", channel: "診所(含門前藥局)", system: "", city: "新北市板橋區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 0, 1971, null, null, null, null, null, null, null] }
      ] },
      { name: "上友兒童診所", channel: "診所(含門前藥局)", system: "", city: "新北市新莊區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 257, 257, null, null, null, null, null, null, null] }
      ] },
      { name: "張振榕診所", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 76, 76, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "殷珮禎", group: "Div3", customers: [
      { name: "秝安藥局(石牌鄭身心診所)", channel: "診所(含門前藥局)", system: "", city: "台北市北投區", products: [
          { series: "Ritalin LA", yearTarget: 920099, monthly: [150302, 83178, 84211, 70840, 152812, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 416784, monthly: [68571, 82286, 82286, 164572, 164572, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 376928, monthly: [50477, 17144, 39048, 21715, 49620, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 121349, monthly: [53334, 53333, 0, 53333, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [4286, 11809, 14286, 25619, 22857, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [22286, 22286, 22286, 7429, 1486, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 17997, monthly: [0, 0, 2666, 0, 2667, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 6115, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3950, monthly: [3162, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 1524, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "正傑健保藥局(心禾診所)", channel: "診所(含門前藥局)", system: "", city: "台北市中山區", products: [
          { series: "Lote", yearTarget: 781326, monthly: [57333, 40952, 73712, 40952, 73713, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 499691, monthly: [16091, 60817, 31445, 110985, 16092, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [11400, 0, 11400, 11400, 4560, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2286, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北辰身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "台北市北投區", products: [
          { series: "Ritalin LA", yearTarget: 497199, monthly: [36287, 34337, 50897, 27069, 63691, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 129493, monthly: [6080, 9867, 8373, 9547, 9547, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 126305, monthly: [5143, 9667, 8000, 6619, 16428, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 116896, monthly: [16857, 0, 16857, 0, 15714, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [42857, 14286, 25000, 8572, 6429, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 15240, monthly: [0, 0, 3387, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3905, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北新醫院", channel: "地區醫院", system: "", city: "新北市淡水區", products: [
          { series: "Exprexa", yearTarget: 425404, monthly: [53872, 0, 26936, 26936, 40768, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 162030, monthly: [28181, 0, 7046, 6907, 27627, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 71804, monthly: [12381, 0, 6190, 0, 12381, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 49557, monthly: [11429, 0, 3048, 3048, 7619, null, null, null, null, null, null, null] },
          { series: "Bisadyl Supp", yearTarget: 999, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "士林身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "台北市士林區", products: [
          { series: "Ritalin LA", yearTarget: 460444, monthly: [0, 54360, 0, 105257, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 59015, monthly: [0, 9571, 0, 15238, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [42857, 0, 8571, 3571, 2143, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 37235, monthly: [14187, 0, 7093, 0, 10400, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 19670, monthly: [0, 16857, 0, 33714, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 9002, monthly: [0, 0, 0, 2571, 0, null, null, null, null, null, null, null] },
          { series: "Bisadyl", yearTarget: 0, monthly: [0, 0, 0, 1715, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "活力藥局(悠活精神科診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "新北市土城區", products: [
          { series: "Ritalin LA", yearTarget: 491612, monthly: [48369, 44155, 21072, 23083, 74931, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 58400, monthly: [8000, 8000, 0, 8000, 8000, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7514, 7514, 3757, 3757, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "社子安心診所", channel: "診所(含門前藥局)", system: "", city: "台北市士林區", products: [
          { series: "Lote", yearTarget: 244160, monthly: [43619, 0, 0, 43619, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 83760, monthly: [21047, 0, 8619, 0, 20333, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 81259, monthly: [16417, 0, 16417, 0, 16417, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 32900, monthly: [10667, 0, 0, 6107, 1947, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 29601, monthly: [8571, 0, 0, 0, 12571, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [10714, 0, 2857, 1429, 5000, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 10080, monthly: [3371, 0, 2248, 0, 2248, null, null, null, null, null, null, null] }
      ] },
      { name: "芸湖藥局(胡耿豪身心精神科診所)", channel: "診所(含門前藥局)", system: "胡耿豪", city: "基隆市安樂區", products: [
          { series: "Ezole", yearTarget: 123640, monthly: [12924, 0, 12486, 24971, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 99321, monthly: [4857, 0, 0, 36571, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 96874, monthly: [0, 13547, 0, 22134, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 78540, monthly: [11905, 0, 0, 11905, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 23400, monthly: [0, 0, 0, 5181, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [3085, 0, 3857, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "天晴診所", channel: "診所(含門前藥局)", system: "", city: "台北市士林區", products: [
          { series: "Ritalin LA", yearTarget: 403528, monthly: [18124, 18124, 0, 48740, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 44228, 0, 9582, 8846, null, null, null, null, null, null, null] }
      ] },
      { name: "開欣藥局(開馨診所)", channel: "診所(含門前藥局)", system: "", city: "台北市南港區", products: [
          { series: "Ritalin LA", yearTarget: 192743, monthly: [19457, 0, 4600, 40686, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 81982, monthly: [15239, 0, 0, 32000, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [13143, 6571, 13143, 7885, 1971, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 31441, monthly: [0, 0, 0, 10933, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 17600, monthly: [3314, 0, 3314, 0, 4419, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 4665, monthly: [0, 1333, 0, 0, 1280, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 0, 515, null, null, null, null, null, null, null] }
      ] },
      { name: "新田身心診所", channel: "診所(含門前藥局)", system: "新田", city: "台北市中山區", products: [
          { series: "Ritalin LA", yearTarget: 332426, monthly: [35648, 8587, 35392, 27935, 34595, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7429, 7429, 3714, 1486, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 0, 625, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 514, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "進安身心診所", channel: "診所(含門前藥局)", system: "進安", city: "台北市中山區", products: [
          { series: "Ezole", yearTarget: 162717, monthly: [14858, 0, 14856, 14858, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 87033, monthly: [0, 10391, 0, 6428, 7715, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 35350, monthly: [0, 8484, 5657, 4242, 4242, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 7143, 3571, 3571, 2858, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 19460, monthly: [2780, 1391, 1391, 1391, 2781, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 4665, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 1658, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "鈺璽診所", channel: "診所(含門前藥局)", system: "", city: "台北市中山區", products: [
          { series: "Ritalin LA", yearTarget: 114027, monthly: [10771, 7180, 9466, 9305, 3591, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 100221, monthly: [0, 0, 1454, 9823, 15583, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 57808, monthly: [7314, 14629, 0, 18287, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 25715, monthly: [5562, 0, 2781, 5334, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林威廷身心精神科診所", channel: "診所(含門前藥局)", system: "林威廷", city: "台北市內湖區", products: [
          { series: "Lote", yearTarget: 148709, monthly: [30714, 0, 36857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 47940, monthly: [8458, 0, 0, 11276, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 38030, monthly: [0, 9524, 4286, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 14286, 0, 7143, 2143, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 18272, monthly: [0, 4571, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "力賀藥局(維品身心診所)", channel: "診所(含門前藥局)", system: "張維仁", city: "新北市中和區", products: [
          { series: "Ritalin LA", yearTarget: 126897, monthly: [32375, 16165, 5918, 21433, 11662, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 43525, monthly: [2826, 7773, 7067, 2826, 11307, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 27874, monthly: [3040, 3040, 3040, 3040, 3040, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [5428, 3571, 1429, 3929, 2857, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 10286, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 1548, 0, 1858, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 2857, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1715, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "健康晨光藥局(合康診所)", channel: "診所(含門前藥局)", system: "", city: "新北市汐止區", products: [
          { series: "Ritalin LA", yearTarget: 176412, monthly: [0, 18289, 24025, 37646, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7371, 7371, 0, 5160, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 5519, monthly: [0, 2454, 0, 6134, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 3330, monthly: [0, 3333, 0, 0, 3333, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 0, 3239, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 3003, monthly: [24286, 0, 0, 0, 20000, null, null, null, null, null, null, null] }
      ] },
      { name: "日揚藥局(新光家醫科診所)", channel: "診所(含門前藥局)", system: "", city: "新北市淡水區", products: [
          { series: "Exelon Patch", yearTarget: 140773, monthly: [4543, 3114, 9087, 16245, 8912, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 128860, monthly: [0, 2837, 11349, 2837, 12769, null, null, null, null, null, null, null] }
      ] },
      { name: "心晴藥局(心築身心診所)", channel: "診所(含門前藥局)", system: "", city: "台北市士林區", products: [
          { series: "Ritalin LA", yearTarget: 79480, monthly: [23126, 16450, 12750, 7715, 9176, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 56502, monthly: [9857, 0, 0, 9429, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7428, 0, 7429, 7428, 2229, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 23660, monthly: [3143, 1858, 3143, 1048, 2095, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 9520, monthly: [0, 0, 1191, 1190, 1191, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 7425, monthly: [0, 0, 1353, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [515, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "双悅診所", channel: "診所(含門前藥局)", system: "暘基", city: "台北市松山區", products: [
          { series: "Ritalin LA", yearTarget: 179493, monthly: [30548, 0, 30549, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 20378, monthly: [6108, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [14286, 0, 3571, 1429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "光點延新藥局(延壽耳鼻喉科診所)", channel: "診所(含門前藥局)", system: "", city: "台北市松山區", products: [
          { series: "Exelon", yearTarget: 95964, monthly: [8161, 1484, 0, 8162, 10386, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 66804, monthly: [6057, 3029, 0, 3114, 1514, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 23542, monthly: [5682, 4059, 0, 4059, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [3152, 328, 1045, 2823, 657, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1114, 1189, 0, 595, 1486, null, null, null, null, null, null, null] }
      ] },
      { name: "心悅身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "新北市淡水區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [35000, 0, 52500, 0, 7000, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 32485, monthly: [14765, 22148, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 28728, monthly: [5095, 6380, 0, 3976, 1524, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 25852, monthly: [6743, 6743, 0, 0, 2248, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1240, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安大心身心科診所", channel: "診所(含門前藥局)", system: "", city: "基隆市中正區", products: [
          { series: "Ritalin LA", yearTarget: 88702, monthly: [24726, 0, 0, 0, 19567, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 31900, monthly: [7048, 0, 12762, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 24595, monthly: [13600, 0, 0, 6186, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3100, monthly: [2477, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2331, 0, 0, 0, 729, null, null, null, null, null, null, null] },
          { series: "Bisadyl", yearTarget: 0, monthly: [2762, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 1714, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "一森診所", channel: "診所(含門前藥局)", system: "", city: "台北市士林區", products: [
          { series: "Exelon Patch", yearTarget: 70996, monthly: [0, 11726, 9142, 8571, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 47082, monthly: [0, 7296, 10714, 4053, 10134, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [12876, 0, 0, 22533, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 514, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "廣福中西藥局(振芝心身醫學診所)", channel: "診所(含門前藥局)", system: "振芝", city: "台北市中山區", products: [
          { series: "Ezole", yearTarget: 109490, monthly: [0, 0, 0, 26000, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 24310, monthly: [5714, 0, 0, 5714, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 17332, monthly: [0, 0, 7429, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1566, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "張國榮身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市士林區", products: [
          { series: "Ritalin LA", yearTarget: 68646, monthly: [23143, 0, 0, 23143, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 33720, monthly: [11239, 0, 0, 11239, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 7143, 0, 1429, 714, null, null, null, null, null, null, null] }
      ] },
      { name: "李政洋身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市松山區", products: [
          { series: "Lote", yearTarget: 38657, monthly: [0, 0, 6233, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 37400, monthly: [0, 4381, 4381, 4381, 4381, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 25290, monthly: [0, 0, 5619, 5619, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 7429, 5200, 2972, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 11200, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3800, monthly: [0, 0, 4571, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "鴻辰藥局(西華診所)", channel: "診所(含門前藥局)", system: "", city: "基隆市七堵區", products: [
          { series: "Exelon", yearTarget: 58032, monthly: [6981, 0, 0, 6981, 6981, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 54296, monthly: [7795, 9353, 3896, 0, 4676, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 0, monthly: [0, 0, 0, 2874, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 743, null, null, null, null, null, null, null] }
      ] },
      { name: "澄心診所", channel: "診所(含門前藥局)", system: "", city: "台北市中山區", products: [
          { series: "Epine", yearTarget: 83620, monthly: [12619, 0, 7429, 13571, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [6011, 0, 9017, 5260, 1503, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 10861, monthly: [0, 0, 1266, 0, 3720, null, null, null, null, null, null, null] }
      ] },
      { name: "心湖身心醫學診所", channel: "診所(含門前藥局)", system: "胡耿豪", city: "新北市汐止區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [18515, 0, 15428, 30857, 77143, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 0, 4000, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1114, 743, 1486, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [1524, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "唯康藥局(樂康診所)(佑家診所)", channel: "診所(含門前藥局)", system: "", city: "台北市松山區", products: [
          { series: "Exelon", yearTarget: 82796, monthly: [5934, 2967, 8901, 4452, 11870, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 10410, monthly: [0, 0, 0, 1735, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [1864, 0, 1864, 0, 2483, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 3006, 0, 0, 1503, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 1733, 0, 0, 1734, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 3098, monthly: [0, 1527, 0, 0, 1439, null, null, null, null, null, null, null] }
      ] },
      { name: "得心身醫學台北總診所", channel: "診所(含門前藥局)", system: "得立", city: "台北市中山區", products: [
          { series: "Ritalin LA", yearTarget: 38736, monthly: [0, 0, 2421, 1543, 1543, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 33540, monthly: [3348, 2423, 2424, 3348, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 22122, monthly: [1227, 2457, 1228, 1228, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 15995, monthly: [915, 915, 915, 915, 915, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 7253, monthly: [2640, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安泰大藥局(悅安身心診所)", channel: "診所(含門前藥局)", system: "", city: "新北市中和區", products: [
          { series: "Lote", yearTarget: 58186, monthly: [3714, 3714, 1238, 2476, 6190, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 23130, monthly: [5429, 1714, 1714, 3523, 1714, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 18676, monthly: [0, 3333, 0, 3334, 3334, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [1486, 1188, 0, 743, 371, null, null, null, null, null, null, null] }
      ] },
      { name: "鄭紹沂診所", channel: "診所(含門前藥局)", system: "", city: "台北市北投區", products: [
          { series: "Exelon", yearTarget: 92459, monthly: [21280, 0, 0, 21280, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [4286, 0, 0, 4286, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "奇岩身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "台北市北投區", products: [
          { series: "Ritalin LA", yearTarget: 43540, monthly: [5250, 3036, 0, 0, 3692, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 29425, monthly: [0, 0, 0, 0, 26667, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7000, 0, 7000, 0, 2100, null, null, null, null, null, null, null] }
      ] },
      { name: "內湖身心精神科診所", channel: "診所(含門前藥局)", system: "林威廷", city: "台北市內湖區", products: [
          { series: "Epine", yearTarget: 31370, monthly: [2857, 0, 0, 11906, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 7143, 7143, 5357, 1428, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 14748, monthly: [2457, 0, 0, 3685, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 9870, monthly: [0, 0, 0, 4228, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "劉宗憲身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市松山區", products: [
          { series: "Exelon", yearTarget: 16478, monthly: [2094, 3490, 4189, 0, 5585, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 15330, monthly: [4381, 4381, 0, 4382, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 12855, monthly: [0, 0, 0, 4285, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 9256, monthly: [1543, 1543, 1543, 6170, 3857, null, null, null, null, null, null, null] }
      ] },
      { name: "康寧醫院", channel: "地區醫院", system: "", city: "台北市內湖區", products: [
          { series: "Bisadyl Supp", yearTarget: 57609, monthly: [3333, 5000, 4000, 5000, 3333, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 10732, monthly: [1306, 1306, 1306, 3267, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "心心藥局(李文珍小兒科診所)", channel: "診所(含門前藥局)", system: "", city: "台北市南港區", products: [
          { series: "Ritalin LA", yearTarget: 36694, monthly: [0, 3883, 6963, 3883, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 9888, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 9091, monthly: [6098, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 3286, 1971, 0, 2472, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 0, monthly: [4347, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 743, 1114, 1486, 743, null, null, null, null, null, null, null] }
      ] },
      { name: "南京春天藥局(昱捷診所)", channel: "診所(含門前藥局)", system: "", city: "台北市松山區", products: [
          { series: "Ritalin LA", yearTarget: 43548, monthly: [10069, 0, 13594, 10069, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [4337, 4337, 2891, 2891, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "福和身心診所", channel: "診所(含門前藥局)", system: "楊思亮", city: "新北市中和區", products: [
          { series: "Ezole", yearTarget: 28474, monthly: [6191, 0, 6191, 6191, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 15700, monthly: [0, 3143, 0, 1500, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 7600, 0, 0, 1520, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5895, monthly: [0, 1314, 1314, 1314, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 1543, 0, 1543, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "培靈醫院", channel: "地區醫院", system: "", city: "台北市松山區", products: [
          { series: "Abimay", yearTarget: 27199, monthly: [4533, 0, 4533, 0, 4533, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 15341, monthly: [3333, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 8743, monthly: [2915, 0, 0, 0, 4337, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [2583, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "悠悅藥局(樂活精神科診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "新北市中和區", products: [
          { series: "Ritalin LA", yearTarget: 29335, monthly: [7695, 0, 0, 7695, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7514, 0, 0, 6011, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 8000, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "沐光藥局(沐陽身心科診所)", channel: "診所(含門前藥局)", system: "", city: "新北市永和區", products: [
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 0, 14509, 18542, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 7307, 3996, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 0, monthly: [0, 0, 0, 6258, 2397, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 2880, 1440, null, null, null, null, null, null, null] }
      ] },
      { name: "安兒康小兒專科診所", channel: "診所(含門前藥局)", system: "", city: "台北市松山區", products: [
          { series: "Ritalin LA", yearTarget: 42582, monthly: [0, 3960, 0, 0, 6600, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1094, 292, 0, 365, null, null, null, null, null, null, null] }
      ] },
      { name: "合康身心診所", channel: "診所(含門前藥局)", system: "張邦彥", city: "新北市土城區", products: [
          { series: "Ezole", yearTarget: 19693, monthly: [16762, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 9331, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [2857, 2857, 0, 1429, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "士林抱抱心身醫學診所", channel: "診所(含門前藥局)", system: "抱抱", city: "台北市士林區", products: [
          { series: "Epine", yearTarget: 19260, monthly: [4190, 0, 1048, 0, 3143, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 10523, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5985, monthly: [0, 0, 1333, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 1333, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 142, 0, 1071, null, null, null, null, null, null, null] }
      ] },
      { name: "抱抱心身醫學診所", channel: "診所(含門前藥局)", system: "抱抱", city: "台北市內湖區", products: [
          { series: "Ezole", yearTarget: 21665, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 16545, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 4655, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 1333, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "詠悅藥局(永平身心診所)", channel: "診所(含門前藥局)", system: "", city: "新北市永和區", products: [
          { series: "Ezole", yearTarget: 28415, monthly: [0, 0, 0, 0, 7714, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 1474, 1474, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 0, 2857, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 2219, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "成合藥局(蘇年興診所)", channel: "診所(含門前藥局)", system: "1B", city: "新北市中和區", products: [
          { series: "Exelon", yearTarget: 12333, monthly: [1366, 0, 1366, 5467, 2734, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 4660, monthly: [4660, 0, 0, 2330, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 743, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 706, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "天母康健身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市士林區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [8400, 0, 8400, 0, 1400, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 10560, monthly: [0, 0, 2640, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 0, 0, 2952, null, null, null, null, null, null, null] }
      ] },
      { name: "明心診所", channel: "診所(含門前藥局)", system: "", city: "新北市永和區", products: [
          { series: "Fute", yearTarget: 14021, monthly: [0, 0, 0, 0, 2667, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 13849, monthly: [0, 0, 0, 0, 3609, null, null, null, null, null, null, null] }
      ] },
      { name: "慶生診所", channel: "診所(含門前藥局)", system: "", city: "台北市中山區", products: [
          { series: "Epine", yearTarget: 23855, monthly: [0, 0, 9047, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "謝宏杰身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市汐止區", products: [
          { series: "Ritalin LA", yearTarget: 29330, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "心禾藥局(詠欣精神科診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "基隆市信義區", products: [
          { series: "Epine", yearTarget: 21547, monthly: [0, 3714, 0, 3714, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "黃偉俐診所", channel: "診所(含門前藥局)", system: "偉俐", city: "台北市中山區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 12331, 0, 1486, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 9198, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [5714, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "永康身心診所(土城)", channel: "診所(含門前藥局)", system: "楊思亮", city: "新北市土城區", products: [
          { series: "Fute", yearTarget: 18995, monthly: [7886, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "靜萱身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市松山區", products: [
          { series: "Abimay", yearTarget: 0, monthly: [0, 0, 0, 3067, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [858, 0, 0, 1071, 1071, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1710, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 1333, monthly: [0, 0, 1333, 2667, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 874, monthly: [0, 0, 1316, 441, 1768, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 771, monthly: [1542, 0, 0, 772, 6170, null, null, null, null, null, null, null] }
      ] },
      { name: "祥新藥局(仁祥診所)", channel: "診所(含門前藥局)", system: "佳醫集康", city: "基隆市仁愛區", products: [
          { series: "Tonex", yearTarget: 0, monthly: [7294, 0, 0, 7294, 3647, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3325, monthly: [1334, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "良哲藥局(常哲診所)", channel: "診所(含門前藥局)", system: "", city: "台北市北投區", products: [
          { series: "Exelon Patch", yearTarget: 9342, monthly: [0, 0, 0, 7786, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [311, 0, 311, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 0, 0, 519, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "偉俐民生診所", channel: "診所(含門前藥局)", system: "偉俐", city: "台北市中山區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [14857, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "好甘心診所", channel: "診所(含門前藥局)", system: "好心肝", city: "台北市中正區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [6571, 0, 6571, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "得悅身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市松山區", products: [
          { series: "Ezole", yearTarget: 0, monthly: [0, 6667, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 1715, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 0, monthly: [0, 1524, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 0, monthly: [0, 1284, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 1143, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "惠生診所", channel: "診所(含門前藥局)", system: "", city: "新北市永和區", products: [
          { series: "Epine", yearTarget: 9894, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1145, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "辰寧身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "台北市北投區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 8858, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 0, 1143, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "寧靜海診所", channel: "診所(含門前藥局)", system: "", city: "新北市中和區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 7566, 0, 1513, null, null, null, null, null, null, null] }
      ] },
      { name: "中崙藥局(中崙國際診所)", channel: "診所(含門前藥局)", system: "", city: "台北市中山區", products: [
          { series: "Ezole", yearTarget: 2855, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 1620, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 1520, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 1280, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 1238, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "永康身心診所(台北)", channel: "診所(含門前藥局)", system: "楊思亮", city: "台北市中正區", products: [
          { series: "Fute", yearTarget: 4585, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3100, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "基隆心身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "基隆市仁愛區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [2800, 0, 700, 0, 700, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 1335, monthly: [0, 0, 1334, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "北投骨科診所", channel: "診所(含門前藥局)", system: "", city: "台北市北投區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 6571, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "陳珀勳診所", channel: "診所(含門前藥局)", system: "", city: "台北市北投區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 0, 0, 4929, null, null, null, null, null, null, null] }
      ] },
      { name: "怡人藥局(舒眠身心醫學診所)", channel: "診所(含門前藥局)", system: "", city: "台北市北投區", products: [
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 0, 0, 3239, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 0, 0, 1500, null, null, null, null, null, null, null] }
      ] },
      { name: "小倉藥局(大昌診所)", channel: "診所(含門前藥局)", system: "", city: "新北市土城區", products: [
          { series: "Tone", yearTarget: 0, monthly: [3714, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "全家健康藥局(頂埔中心診所)", channel: "診所(含門前藥局)", system: "", city: "新北市土城區", products: [
          { series: "Epine", yearTarget: 1380, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "洪聖哲", group: "Div3", customers: [
      { name: "翰生藥局(翰生診所)", channel: "診所(含門前藥局)", system: "", city: "新北市板橋區", products: [
          { series: "Ritalin LA", yearTarget: 145625, monthly: [94263, 0, 94263, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [30171, 0, 30171, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5577, monthly: [0, 0, 5143, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3982, monthly: [6381, 0, 1238, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "夏凱納診所", channel: "診所(含門前藥局)", system: "", city: "台北市大同區", products: [
          { series: "Ritalin LA", yearTarget: 97327, monthly: [40971, 52002, 8720, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 49043, monthly: [38429, 0, 51239, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 10338, monthly: [5809, 2714, 4261, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 4997, monthly: [3572, 0, 7143, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 3011, monthly: [0, 2667, 8000, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [1440, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "昀樂身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市新莊區", products: [
          { series: "Ritalin LA", yearTarget: 19589, monthly: [152229, 0, 46285, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [11886, 0, 7429, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 3235, monthly: [0, 0, 7857, 0, 0, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 1779, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 1753, monthly: [3143, 0, 3333, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 917, monthly: [0, 8666, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 884, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 341, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "安興精神科診所", channel: "診所(含門前藥局)", system: "安興", city: "新北市三重區", products: [
          { series: "Ritalin LA", yearTarget: 44400, monthly: [37283, 0, 8800, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 19900, monthly: [7429, 0, 24286, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 14208, monthly: [16191, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 13432, monthly: [0, 11600, 4771, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 7033, monthly: [12381, 0, 4953, 0, 0, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 2744, monthly: [0, 6347, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "仁濟新莊", channel: "地區醫院", system: "仁濟", city: "新北市新莊區", products: [
          { series: "Exprexa", yearTarget: 59625, monthly: [36373, 0, 9920, 0, 0, null, null, null, null, null, null, null] },
          { series: "Cospirit", yearTarget: 29728, monthly: [3733, 0, 14000, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 27214, monthly: [3095, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 23575, monthly: [0, 0, 12802, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 6369, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "晴美身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市新莊區", products: [
          { series: "Ritalin LA", yearTarget: 107637, monthly: [0, 0, 84000, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [28457, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2398, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 307, monthly: [1429, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "祐康藥局(養心診所)", channel: "診所(含門前藥局)", system: "", city: "新北市鶯歌區", products: [
          { series: "Lote", yearTarget: 64650, monthly: [22762, 22762, 68285, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 15444, monthly: [18286, 0, 9001, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "心世界身心精神科診所", channel: "診所(含門前藥局)", system: "", city: "新北市板橋區", products: [
          { series: "Lote", yearTarget: 34571, monthly: [0, 17333, 26666, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 14282, monthly: [0, 9638, 12571, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 10547, monthly: [0, 12191, 12191, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 10120, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5429, monthly: [0, 3143, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 4480, monthly: [0, 3993, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "心安藥局(育心身心精神科診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "新北市新莊區", products: [
          { series: "Ritalin LA", yearTarget: 48029, monthly: [19603, 0, 23451, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 34869, monthly: [0, 18572, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7514, 0, 3757, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 9089, monthly: [4000, 4000, 4000, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "和沛身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "台北市文山區", products: [
          { series: "Ritalin LA", yearTarget: 49715, monthly: [48023, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 21923, monthly: [0, 0, 16619, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 10255, monthly: [9686, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 7143, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 3810, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "欣悅藥局(美麗心成人兒童精神科診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "新北市三重區", products: [
          { series: "Ritalin LA", yearTarget: 61741, monthly: [0, 31145, 35360, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 15907, monthly: [8000, 0, 8000, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂心藥局(順心診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "新北市板橋區", products: [
          { series: "Ritalin LA", yearTarget: 34470, monthly: [6322, 27297, 30778, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7514, 6011, 3757, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 13764, monthly: [7429, 0, 7428, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 6916, monthly: [4000, 2400, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "福田診所", channel: "診所(含門前藥局)", system: "", city: "新北市蘆洲區", products: [
          { series: "Ritalin LA", yearTarget: 57056, monthly: [38443, 0, 38443, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 7143, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "芯安藥局(欣寧身心診所)", channel: "診所(含門前藥局)", system: "", city: "新北市三峽區", products: [
          { series: "Ritalin LA", yearTarget: 41500, monthly: [23143, 23143, 46285, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "心翔身心精神科診所", channel: "診所(含門前藥局)", system: "林威廷", city: "新北市新莊區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [17857, 0, 3571, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 12142, monthly: [9829, 0, 24572, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 11967, monthly: [14286, 0, 21905, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3657, monthly: [4228, 0, 4228, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 1975, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "木柵身心診所", channel: "診所(含門前藥局)", system: "", city: "台北市文山區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 73029, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 41303, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 8165, monthly: [3429, 3429, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "欣泉身心診所", channel: "診所(含門前藥局)", system: "張邦彥", city: "新北市林口區", products: [
          { series: "Lote", yearTarget: 25682, monthly: [13333, 0, 13333, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 19146, monthly: [16762, 0, 16761, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 10152, monthly: [0, 0, 6712, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [514, 0, 1543, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 1429, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "藥濟藥局(欣慈診所)", channel: "診所(含門前藥局)", system: "崇文", city: "新北市鶯歌區", products: [
          { series: "Dalmadorm", yearTarget: 0, monthly: [7048, 7048, 10571, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 21218, monthly: [8150, 5930, 7715, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 12053, monthly: [6095, 0, 3048, 0, 0, null, null, null, null, null, null, null] },
          { series: "Slatone", yearTarget: 9078, monthly: [4191, 0, 4190, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 0, monthly: [0, 0, 8095, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 5571, monthly: [0, 2191, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "龍霖身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市林口區", products: [
          { series: "Ritalin LA", yearTarget: 55988, monthly: [22637, 0, 22637, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 7429, 2971, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "合誼身心診所", channel: "診所(含門前藥局)", system: "進安", city: "新北市蘆洲區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [3571, 7143, 7143, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 14865, monthly: [7429, 0, 14858, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3605, monthly: [4172, 0, 2781, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 2404, monthly: [7554, 0, 11481, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 1746, monthly: [1414, 2829, 1414, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "福全身心科診所", channel: "診所(含門前藥局)", system: "", city: "台北市萬華區", products: [
          { series: "Epine", yearTarget: 18226, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 18162, monthly: [0, 0, 28286, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 7143, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 3191, monthly: [0, 7619, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "西園醫院", channel: "地區醫院", system: "", city: "台北市萬華區", products: [
          { series: "Lote", yearTarget: 46320, monthly: [14095, 0, 21143, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "恆友精神科診所", channel: "診所(含門前藥局)", system: "安興", city: "新北市新莊區", products: [
          { series: "Ritalin LA", yearTarget: 27243, monthly: [0, 7852, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 4586, monthly: [0, 3715, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3848, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 3319, monthly: [0, 7333, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2202, monthly: [7429, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Abimay", yearTarget: 784, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "治聿身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Ritalin LA", yearTarget: 17701, monthly: [7714, 0, 7715, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 7669, monthly: [4723, 0, 4723, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [5100, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3343, 0, 743, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon", yearTarget: 3572, monthly: [2152, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新店北新身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [7143, 7143, 7143, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 15449, monthly: [6212, 3106, 3106, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 12427, monthly: [0, 2648, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "放開心身心診所", channel: "診所(含門前藥局)", system: "好心情", city: "新北市板橋區", products: [
          { series: "Fute", yearTarget: 19429, monthly: [27428, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7429, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2569, monthly: [0, 0, 2400, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1943, monthly: [0, 0, 1429, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 567, monthly: [1048, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "均樂身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Ritalin LA", yearTarget: 22440, monthly: [31429, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 7143, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "佳醫-新泰醫院", channel: "地區醫院", system: "佳醫集康", city: "新北市新莊區", products: [
          { series: "Epram", yearTarget: 21302, monthly: [10572, 0, 3524, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [5051, 3156, 3157, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tonex", yearTarget: 6306, monthly: [3648, 3648, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "喜樂藥局(喜悅診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "新北市樹林區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [7514, 0, 7514, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 9485, monthly: [4000, 0, 8000, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 3972, monthly: [3078, 7145, 7988, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新店開心診所", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Ritalin LA", yearTarget: 42475, monthly: [0, 0, 15714, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "楊思亮診所", channel: "診所(含門前藥局)", system: "楊思亮", city: "新北市板橋區", products: [
          { series: "Ezole", yearTarget: 12537, monthly: [12381, 0, 12381, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5429, monthly: [7858, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 3398, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 307, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "萬華維沐身心診所", channel: "診所(含門前藥局)", system: "張維仁", city: "台北市萬華區", products: [
          { series: "Ritalin LA", yearTarget: 20520, monthly: [13105, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 8744, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 2857, 1428, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 1753, monthly: [0, 3040, 2027, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "樂開懷健保藥局(開心生活診所)", channel: "診所(含門前藥局)", system: "開心房", city: "新北市蘆洲區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [7428, 0, 7429, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 2236, monthly: [7972, 0, 15942, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2164, monthly: [7620, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "柏康藥局(丁良文內科診所)", channel: "診所(含門前藥局)", system: "", city: "新北市蘆洲區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [2590, 16429, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 10737, monthly: [6286, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5447, monthly: [0, 5238, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1990, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "誠心身心醫學診所", channel: "診所(含門前藥局)", system: "", city: "台北市文山區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [7429, 7429, 3714, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 11091, monthly: [15531, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "向陽澄思診所", channel: "診所(含門前藥局)", system: "張邦彥", city: "新北市板橋區", products: [
          { series: "Ezole", yearTarget: 9521, monthly: [0, 0, 16762, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 8557, monthly: [0, 0, 2685, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [3571, 0, 2857, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "楊孟達身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "新北市板橋區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 10500, 3500, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 10543, monthly: [0, 0, 7383, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 2637, monthly: [5333, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 705, monthly: [0, 0, 1143, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新信義藥局(欣彥身心診所)", channel: "診所(含門前藥局)", system: "", city: "新北市蘆洲區", products: [
          { series: "Epine", yearTarget: 8017, monthly: [3429, 0, 3429, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 5039, monthly: [0, 0, 9000, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 2317, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1690, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 1542, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "蘇宗偉身心診所", channel: "診所(含門前藥局)", system: "", city: "桃園市龜山區", products: [
          { series: "Epine", yearTarget: 9874, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exprexa", yearTarget: 7566, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 1734, monthly: [11714, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "興隆藥局(方舟身心診所)", channel: "診所(含門前藥局)", system: "烱鳴", city: "台北市文山區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [6763, 0, 4508, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 5229, monthly: [0, 3848, 7695, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [0, 0, 1020, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "宏通診所", channel: "診所(含門前藥局)", system: "", city: "新北市三重區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [6571, 0, 6571, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 5207, monthly: [5238, 0, 3809, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "詠悅藥局(永平身心診所)", channel: "診所(含門前藥局)", system: "", city: "新北市永和區", products: [
          { series: "Ezole", yearTarget: 9321, monthly: [0, 5904, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [7371, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 690, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "陳信任精神科診所", channel: "診所(含門前藥局)", system: "", city: "新北市三重區", products: [
          { series: "Epine", yearTarget: 11604, monthly: [8096, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1162, monthly: [2279, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "仁濟台北", channel: "地區醫院", system: "仁濟", city: "台北市萬華區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [3334, 0, 6667, 0, 0, null, null, null, null, null, null, null] },
          { series: "Trileptal", yearTarget: 0, monthly: [0, 0, 6571, 0, 0, null, null, null, null, null, null, null] },
          { series: "Exelon Patch", yearTarget: 5791, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "王湘琦身心診所", channel: "診所(含門前藥局)", system: "", city: "新北市三峽區", products: [
          { series: "Fute", yearTarget: 3335, monthly: [4057, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 2869, monthly: [2524, 0, 2524, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 1378, monthly: [3715, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "好心情身心診所", channel: "診所(含門前藥局)", system: "好心情", city: "新北市新莊區", products: [
          { series: "Fute", yearTarget: 10701, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 4457, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 1482, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 1429, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 260, monthly: [0, 0, 1047, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林口抱抱心身醫學診所", channel: "診所(含門前藥局)", system: "抱抱", city: "新北市林口區", products: [
          { series: "Ritalin LA", yearTarget: 2448, monthly: [3200, 0, 4800, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [714, 714, 714, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 1071, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 778, monthly: [0, 0, 2095, 0, 0, null, null, null, null, null, null, null] },
          { series: "Dalmadorm", yearTarget: 0, monthly: [313, 0, 312, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "中山抱抱心身醫學診所", channel: "診所(含門前藥局)", system: "抱抱", city: "台北市大同區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [4000, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 2628, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [714, 1429, 357, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 612, monthly: [2476, 0, 2476, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "林南診所", channel: "診所(含門前藥局)", system: "", city: "新北市三重區", products: [
          { series: "Ritalin LA", yearTarget: 3761, monthly: [0, 4472, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 3359, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 357, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "溫時光診所", channel: "診所(含門前藥局)", system: "", city: "新北市新莊區", products: [
          { series: "Ezole", yearTarget: 8803, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Tofranil", yearTarget: 0, monthly: [0, 2571, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "壬康診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "新北市蘆洲區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [1050, 0, 2100, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 1694, monthly: [739, 0, 739, 0, 0, null, null, null, null, null, null, null] },
          { series: "Lote", yearTarget: 1666, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ezole", yearTarget: 1318, monthly: [1904, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "萬華醫院", channel: "地區醫院", system: "", city: "台北市萬華區", products: [
          { series: "Exelon Patch", yearTarget: 7821, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "新田身心診所龜山", channel: "診所(含門前藥局)", system: "新田", city: "桃園市龜山區", products: [
          { series: "Ritalin LA", yearTarget: 6491, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "沐林親子診所", channel: "診所(含門前藥局)", system: "", city: "新北市板橋區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [4787, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [148, 1486, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "萬華身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "台北市萬華區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 4900, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 330, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "幸福身心診所", channel: "診所(含門前藥局)", system: "賈惠洲", city: "新北市板橋區", products: [
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 1400, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Fute", yearTarget: 705, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 330, monthly: [0, 1334, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "後埔永江診所", channel: "診所(含門前藥局)", system: "", city: "新北市板橋區", products: [
          { series: "Fute", yearTarget: 3321, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] },
          { series: "Ritalin LA", yearTarget: 339, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "建國六號藥局(蘇景傑診所)", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Tofranil", yearTarget: 0, monthly: [0, 3620, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "正欣藥局(正興診所)", channel: "診所(含門前藥局)", system: "", city: "新北市三重區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [3286, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "倍思特藥局(楊聰才診所)", channel: "診所(含門前藥局)", system: "", city: "新北市新店區", products: [
          { series: "Tone", yearTarget: 0, monthly: [0, 3048, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "里安診所", channel: "診所(含門前藥局)", system: "", city: "新北市板橋區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [0, 1643, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "富康藥局(建順診所)", channel: "診所(含門前藥局)", system: "", city: "台北市大同區", products: [
          { series: "Lote", yearTarget: 1199, monthly: [0, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] },
      { name: "上友兒童診所", channel: "診所(含門前藥局)", system: "", city: "新北市新莊區", products: [
          { series: "Trileptal", yearTarget: 0, monthly: [257, 0, 0, 0, 0, null, null, null, null, null, null, null] }
      ] }
    ] },
    { name: "黃建源", group: "Div3", customers: [
      { name: "橙易藥局(吳四維診所)", channel: "診所(含門前藥局)", system: "", city: "苗栗縣苗栗市", products: [
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 0, 0, 39400, null, null, null, null, null, null, null] },
          { series: "Epine", yearTarget: 0, monthly: [0, 0, 0, 0, 3428, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 0, 0, 2857, null, null, null, null, null, null, null] },
          { series: "Ritalin IR", yearTarget: 0, monthly: [0, 0, 0, 0, 2074, null, null, null, null, null, null, null] }
      ] },
      { name: "快樂藥局(時光身心診所)", channel: "診所(含門前藥局)", system: "", city: "台中市大雅區", products: [
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 0, 0, 26475, null, null, null, null, null, null, null] },
          { series: "Tone", yearTarget: 0, monthly: [0, 0, 0, 0, 1429, null, null, null, null, null, null, null] }
      ] },
      { name: "樂活藥局(心身美診所)", channel: "診所(含門前藥局)", system: "", city: "台中市南屯區", products: [
          { series: "Ritalin LA", yearTarget: 0, monthly: [0, 0, 0, 0, 15691, null, null, null, null, null, null, null] }
      ] },
      { name: "全新生活診所", channel: "診所(含門前藥局)", system: "", city: "台中市西屯區", products: [
          { series: "Lote", yearTarget: 0, monthly: [0, 0, 0, 0, 6571, null, null, null, null, null, null, null] }
      ] }
    ] }
  ]
};
