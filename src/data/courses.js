const courses = {
  SP: { name: 'Əmniyyətli İdarəetmə haqqında Beynəlxalq Məcəllə', hours: 16 },
  SI: { name: 'Gəminin Mühafizəsi üzrə ümumi hazırlıq və təlimat', hours: 8 },
  SH: { name: 'Gəminin Mühafizəsi üzrə müəyyən edilmiş vəzifələrə malik şəxslər', hours: 16 },
  SG: { name: 'Gəmi mühafizəsi üzrə məsul Şəxs', hours: 18 },
  SO: { name: 'Bütün dənizçilər üçün təhlükəsizlik üzrə tanışlıq və təlimat', hours: 80 },
  SW: { name: 'Kapitan Körpüsü Resurslarının İdarə Olunması', hours: 42 },
  SV: { name: 'Gəminin İdarə Olunması və Manevr Edilməsi', hours: 40 },
  SQ: { name: 'Radar, ARPA, körpü komandası və axtarış-xilasetmə (idarəetmə)', hours: 40 },
  SR: { name: 'Radar müşahidəsi və ARPA-nın istismarı (operativ səviyyə)', hours: 98 },
  SZ: { name: 'Elektron Xəritə Displeyi və İnformasiya Sistemlərinin İstismarı', hours: 40 },
  SF: { name: 'Sərnişinlərin, yükün və gövdənin təhlükəsizliyi (Ro-Ro)', hours: 18 },
  SD: { name: 'Sərnişinlərə xidmət göstərən heyət üçün təhlükəsizlik', hours: 8 },
  SC: { name: 'İzdihamın idarə olunması üzrə hazırlıq', hours: 11 },
  SE: { name: 'Böhran zamanı idarəetmə və insan davranışı üzrə hazırlıq', hours: 16 },
  ST: { name: 'Gəmi qaz analizatorları və onların istismarı', hours: 8 },
  SX: { name: 'İnert qaz sistemi', hours: 16 },
  SN: { name: 'Gəmidə ilk tibbi yardım', hours: 34 },
  SM: { name: 'Gəmidə tibbi nəzarət', hours: 47 },
  DQ: { name: 'Qlobal Dəniz Fəlakət və Əmniyyətli Rabitə Sisteminin Operatoru', hours: 110 },
  SA: { name: 'Neft və kimyəvi tankerlərdə yük əməliyyatına dair ilkin hazırlıq', hours: 48 },
  SB: { name: 'Neft tankerlərdə geniş proqram üzrə hazırlıq', hours: 55 },
  AS: { name: 'Kimyəvi tankerlərdə geniş proqram üzrə hazırlıq', hours: 60 },
  SK: { name: 'Təhlükəli və zərərli yüklərin daşınması', hours: 34 },
  ER: { name: 'Maşın şöbəsinin resurslarının idarə olunması', hours: 37 },
  DL: { name: 'Liderlik və heyətlə iş birliyi', hours: 20 },
  SJ: { name: 'Yanğınla mübarizə geniş proqram üzrə', hours: 32 },
  SL: { name: 'Sürətli olmayan xilasedici qayıq üzrə mütəxəssis', hours: 32 },
  SU: { name: 'Sürətli xilasetmə qayıqları üzrə mütəxəssis', hours: 20 },
  WS: { name: 'Liman Vasitələrinin mühafizəsinə məsul şəxsin hazırlığı', hours: 20 },
  RS: { name: 'Reytinq sxemi üzrə hazırlıq', hours: 48 },
};

export const getCourseName = (code) => courses[code]?.name || 'Ad təyin olunmayıb';
export const getCourseHours = (code) => courses[code]?.hours || 0;
export const getCourseInfo = (code) => courses[code] || { name: 'Ad təyin olunmayıb', hours: 0 };
export default courses;
