const courses = {
  SK: { name: 'Gəmi sürücülərinin təkmilləşdirməsi kursu', hours: 160 },
  SH: { name: 'Gəmidə ilk tibbi yardım', hours: 34 },
  SI: { name: 'Sürətli olmayan xilasedici qayıq, növbətçi qayıq və sallar üzrə mütəxəssis', hours: 32 },
  SO: { name: 'Gəmi sürücülərinin təkmilləşdirməsi kursu (idarəetmə səviyyəsi)', hours: 160 },
  SL: { name: 'Təhlükəli və zərərli yüklərin daşınması', hours: 34 },
  SA: { name: 'Gəmi sürücülərinin hazırlanması kursu', hours: 160 },
  SP: { name: 'Maşın şöbəsinin resurslarının idarə olunması', hours: 37 },
  RS: { name: 'Reytinq sxemi üzrə hazırlıq', hours: 48 },
  SG: { name: 'Gəmi mexaniki hazırlıq kursu', hours: 120 },
  SW: { name: 'Dənizçilik təhlükəsizliyi (ISPS)', hours: 16 },
  SV: { name: 'Beynəlxalq Dəniz Təşkilatı (IMO) Model Qanunu', hours: 24 },
  SQ: { name: 'Kimyəvi maddələr daşıyan tankerlərdə yük əməliyyatına dair geniş proqram', hours: 60 },
  SR: { name: 'Gəmi elektriki hazırlıq kursu', hours: 120 },
  SZ: { name: 'Sürətli xilasetmə qayıqlar üzrə mütəxəssis', hours: 20 },
  SF: { name: 'Qaz yandıran mühərrik üzrə hazırlıq', hours: 48 },
  SD: { name: 'ROV pilotu hazırlıq kursu', hours: 8 },
  SC: { name: 'Sualtı əməliyyatlar üzrə hazırlıq', hours: 24 },
  SE: { name: 'STCW-başlanğıc təhlükəsizlik', hours: 16 },
  ST: { name: 'Yanğınsöndürmə hazırlıq kursu', hours: 8 },
  SX: { name: 'Gəmi operatoru hazırlıq kursu', hours: 16 },
  SN: { name: 'Növbə mexaniki hazırlıq kursu', hours: 120 },
  SM: { name: 'Maşinist hazırlıq kursu', hours: 120 },
  DQ: { name: 'Dəniz qazancı hazırlıq kursu', hours: 200 },
  SB: { name: 'Bərə operatoru hazırlıq kursu', hours: 56 },
  AS: { name: 'Aşpaz hazırlıq kursu', hours: 120 },
  ER: { name: 'Elektrik mexaniki hazırlıq kursu', hours: 80 },
  DL: { name: 'Dalğıc hazırlıq kursu', hours: 40 },
  SJ: { name: 'Kran maşinisti hazırlıq kursu', hours: 64 },
  SU: { name: 'Sürətli xilasetmə qayıqlar üzrə mütəxəssis', hours: 20 },
  WS: { name: 'Yükləmə əməliyyatları hazırlıq kursu', hours: 32 },
  XS: { name: 'Xüsusi hazırlıq kursu', hours: 24 },
};

export const getCourseName = (code) => courses[code]?.name || 'Ad təyin olunmayıb';
export const getCourseHours = (code) => courses[code]?.hours || 0;
export const getCourseInfo = (code) => courses[code] || { name: 'Ad təyin olunmayıb', hours: 0 };
export default courses;
