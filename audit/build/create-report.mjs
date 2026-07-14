import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = new URL("../", import.meta.url).pathname;
const checkedAt = "2026-07-14 12:15 +05:00";

const workbook = Workbook.create();

const sheets = [
  {
    name: "Clinic Data",
    columns: ["Поле", "Значение", "Источник", "Статус"],
    rows: [
      ["Название", "Aident — цифровая стоматология", "https://2gis.kz/aktau/firm/70000001028449496", "Подтверждено 2GIS"],
      ["Адрес", "28-й микрорайон, 1, офис 4, 1 этаж, Актау", "https://2gis.kz/aktau/firm/70000001028449496", "Подтверждено 2GIS"],
      ["Телефон", "+7 775 825 52 43", "https://2gis.kz/aktau/firm/70000001028449496", "Подтверждено 2GIS"],
      ["WhatsApp", "https://wa.me/77758255243", "Карточка 2GIS", "Использован как CTA в обновлённой концепции по запросу пользователя; требуется подтверждение владельцем, поскольку bio Instagram ведёт на wa.me/7758255243"],
      ["Instagram", "https://www.instagram.com/aident_clinic/", "https://2gis.kz/aktau/firm/70000001028449496; публичный профиль Instagram", "Подтверждено: карточка 2GIS ссылается на этот профиль; совпадают бренд, логотип и адрес 28-1"],
      ["2GIS", "https://2gis.kz/aktau/firm/70000001028449496", "https://2gis.kz/aktau/firm/70000001028449496", "Подтверждено"],
      ["График", "Не публиковать до подтверждения владельцем", "2GIS и Instagram bio", "РАСХОЖДЕНИЕ: 2GIS указывает более короткий субботний график; Instagram — Пн–Сб 09:00–18:00, перерыв 13:00–14:00"],
      ["Рейтинг", "4.9", "https://2gis.kz/aktau/firm/70000001028449496", "Наблюдение 14.07.2026"],
      ["Отзывы", "Рейтинг 4.9; в концепции использованы два коротких публичных фрагмента", "https://2gis.kz/aktau/firm/70000001028449496/tab/reviews", "Наблюдение 14.07.2026; число оценок сознательно не публикуется и может меняться"],
      ["Услуги", "Изготовление коронок; 3D-моделирование и печать; CAD/CAM, виниры, E-max; интраоральное 3D-сканирование; лечение; протезирование; имплантация", "Instagram bio; https://www.instagram.com/aident_clinic/p/DXdrmjpioFi/", "Подтверждено официальным профилем"],
      ["Языки", "Русский, казахский", "https://2gis.kz/aktau/firm/70000001028449496", "Подтверждено 2GIS"],
      ["Логотип", "Аватар Instagram, 1080×1080", "https://www.instagram.com/aident_clinic/", "Подтверждено; сохранён как media/logo/aident-instagram-avatar.jpg"],
      ["Врачи / цены / гарантии", "Не публиковать", "Публичные источники проверены", "Не подтверждены имена, специализации и согласие на использование фото; раздел врачей исключён"],
    ],
  },
  {
    name: "Website Verification",
    columns: ["Источник", "Проверенный URL", "Результат", "Найденная ссылка", "Тип страницы", "Дата проверки", "Скриншот", "Комментарий", "Итоговый вердикт"],
    rows: [
      ["2GIS", "https://2gis.kz/aktau/firm/70000001028449496", "Отдельный сайт не указан", "Instagram; WhatsApp; телефон", "Каталог / карта", checkedAt, "audit/screenshots/web/2gis-contacts.webp", "На скриншоте видны телефон, Instagram и WhatsApp; отдельное поле «Сайт» отсутствует.", "Пройдено только для 2GIS"],
      ["Instagram", "https://www.instagram.com/aident_clinic/", "Отдельный сайт, мини-сайт, Tilda, Taplink, Linktree и запись не указаны", "Bio: wa.me/7758255243; Threads: @aident_clinic", "Официальная социальная сеть", checkedAt, "audit/screenshots/web/instagram-profile-header.webp", "Профиль открыт в Chrome: совпадают бренд AiDent, аватар-логотип и адрес 28-1 с 2GIS. Единственная внешняя ссылка bio после редиректа ведёт на https://wa.me/7758255243/; видны кнопки «Подписаться» и «Отправить сообщение». 7 highlights перечислены; доступные закреплённые материалы не содержат доменной ссылки. Одна публикация (DXdqtbRisXQ) не загрузилась в UI — это зафиксировано, но не является ссылкой на сайт.", "Пройдено: официальный сайт не обнаружен"],
      ["Независимый поиск", "Запросы: Aident / AiDent / aident_clinic / телефон / адрес", "Официальный домен, Tilda, Taplink, Linktree и сайт записи не найдены", "2GIS; Яндекс Карты; Instagram; каталоги; Enbek", "Поиск / каталоги / соцсети", checkedAt, "Не применимо", "Похожие бренды и неофициальные записи исключены. Результат согласуется с проверкой 2GIS и Instagram.", "Пройдено: официальный сайт не обнаружен"],
      ["Итог", "", "Полноценный официальный сайт не найден — переход к предварительной концепции разрешён.", "", "", checkedAt, "", "Три независимых источника не обнаружили сайт. До согласования владельцем нужно сохранить noindex, nofollow и не использовать спорные WhatsApp/график.", "РАЗРАБОТКА РАЗРЕШЕНА"],
    ],
  },
  {
    name: "Sources and Media",
    columns: ["Тип материала", "Локальное имя файла", "URL оригинального источника", "Где используется", "Автор источника", "Дата проверки", "Комментарий"],
    rows: [
      ["Скриншот контактов", "audit/screenshots/web/2gis-contacts.webp", "https://2gis.kz/aktau/firm/70000001028449496", "Доказательство проверки", "2GIS", checkedAt, "Оптимизированная публичная копия контактов; не является медиа для сайта."],
      ["Скриншот профиля Instagram", "audit/screenshots/web/instagram-profile-header.webp", "https://www.instagram.com/aident_clinic/", "Доказательство проверки", "aident_clinic", checkedAt, "Оптимизированная копия заголовка профиля, логотипа, bio, ссылки WhatsApp и контактной кнопки; персональная навигация пользователя не попала в кадр."],
      ["Логотип", "media/logo/aident-instagram-avatar.jpg", "https://www.instagram.com/aident_clinic/", "Планируется: шапка, favicon, Open Graph", "aident_clinic", checkedAt, "Официальный аватар профиля, публично получен через страницу, 1080×1080; пропорции не изменены."],
      ["Отзывы", "Нет локального медиа", "https://2gis.kz/aktau/firm/70000001028449496/tab/reviews", "Раздел отзывов", "2GIS; публичные авторы отзывов", checkedAt, "Использованы только короткие цитаты: «Золотые руки и трепетное отношение» (Margarita Zholdasova) и «Отличная клиника!» (Exxon M)."],
      ["Карта", "site/assets/media/map/aident-2gis-map.png", "https://beautyshare.2gis.ru/api/v1/image?city=aktau&zoom=17&center=51.164119%2C43.668692&title=Aident&desc=28-%D0%B9%20%D0%BC%D0%B8%D0%BA%D1%80%D0%BE%D1%80%D0%B0%D0%B9%D0%BE%D0%BD%2C%201%2C%20%D0%90%D0%BA%D1%82%D0%B0%D1%83", "Блок контактов; ссылка на карточку 2GIS", "2GIS", checkedAt, "Публичный статичный превью-слой 2GIS с точкой Aident; не используется Google Maps."],
    ],
  },
  {
    name: "Website Production",
    columns: ["Концепция", "Фирменные цвета", "Стек", "Репозиторий", "Публичный сайт", "Хостинг", "Дата деплоя", "Статус"],
    rows: [["Предварительный одностраничный сайт: выбор услуг, о клинике, отзывы 2GIS, карта 2GIS и запись в WhatsApp", "Чёрный/белый логотип; тёплый нейтральный интерфейс", "HTML, CSS, JavaScript", "https://github.com/nuradilabyz/aident-aktau-website-concept", "https://nuradilabyz.github.io/aident-aktau-website-concept/", "GitHub Pages через GitHub Actions", "14.07.2026", "Опубликовано и повторно проверено QA. noindex, nofollow и запрет в robots.txt сохранены. WhatsApp взят из 2GIS и ждёт подтверждения владельцем; график не опубликован."]],
  },
  {
    name: "QA Report",
    columns: ["Устройство", "Разрешение", "Скриншот", "Проверенные элементы", "Найденные ошибки", "Исправления", "Повторная проверка", "Итоговый статус"],
    rows: [
      ["Компьютер", "1440 px", "audit/qa-screenshots/redesign-recheck-1440-full.png; redesign-recheck-1440-contacts-map.png", "Полная страница, hero, услуги, отзывы, контакты и карта; HTTPS; ссылки; console; SEO и OG", "На первом проходе карта 2GIS была чрезмерно затемнена — география читалась плохо", "Ослаблено overlay-затемнение; CTA маршрута получила отдельную контрастную подложку; опубликован новый CSS URL", "Повторно проверено на live-версии ?qa=map-fix-20260714: улицы, точка Aident, адрес и CTA читаемы", "QA PASSED"],
      ["Планшет", "768 px", "audit/qa-screenshots/redesign-recheck-768-full.png", "Полная страница; hero, услуги, отзывы, контакты; отсутствие горизонтальной прокрутки и ошибок консоли", "Нет после исправления", "Не требуется", "На live-версии 753 px контента ≤ 768 px viewport; console []", "QA PASSED"],
      ["Мобильный", "390 px", "audit/qa-screenshots/redesign-recheck-390-full.png; redesign-recheck-390-menu-open.png; redesign-recheck-390-contacts-map.png; redesign-recheck-390-map-lower-sticky.png", "Полная страница, hero, услуги, отзывы, мобильное меню, фиксированная WhatsApp-кнопка, карта и нижняя область", "Нет после исправления", "Не требуется", "На live-версии 375 px контента ≤ 390 px viewport; меню aria-expanded=true; CTA не закрывает кнопку карты", "QA PASSED"],
    ],
  },
  {
    name: "Outreach",
    columns: ["Название", "Номер", "WhatsApp", "Instagram", "Публичная ссылка", "Основное сообщение", "Короткое сообщение", "Сценарий звонка", "Ответ на возражение", "Статус связи"],
    rows: [["Aident", "+7 775 825 52 43", "https://wa.me/77758255243 (требует подтверждения владельцем)", "https://www.instagram.com/aident_clinic/", "https://nuradilabyz.github.io/aident-aktau-website-concept/", "Здравствуйте, команда Aident! Подготовил предварительную концепцию сайта для вашей цифровой стоматологии в Актау — с выбором услуг, короткими отзывами из 2GIS и маршрутом 2GIS: https://nuradilabyz.github.io/aident-aktau-website-concept/ Буду рад показать проект и обсудить правки.", "Здравствуйте! Подготовил концепцию сайта для Aident с услугами, 2GIS и WhatsApp-записью. Посмотреть: https://nuradilabyz.github.io/aident-aktau-website-concept/ Могу коротко показать идею и обсудить правки.", "Здравствуйте! Я подготовил предварительную концепцию сайта Aident на основе 2GIS и официального Instagram. Можно показать ссылку и обсудить, что важно оставить или изменить?", "Понимаю. Проект ни к чему не обязывает: он показывает, как объединить запись в WhatsApp, 2GIS и цифровые направления клиники. Если сейчас неактуально, можно сохранить ссылку на будущее.", "Не отправлять самостоятельно; номер WhatsApp и график требуют подтверждения владельцем"]],
  },
];

const headerFormat = { fill: "#075E54", font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
const titleFormat = { fill: "#E7F4F1", font: { bold: true, color: "#073B35", size: 15 }, horizontalAlignment: "left", verticalAlignment: "center" };

for (const spec of sheets) {
  const sheet = workbook.worksheets.add(spec.name);
  sheet.showGridLines = false;
  const cols = spec.columns.length;
  const lastColumn = String.fromCharCode(64 + cols);
  sheet.getRange(`A1:${lastColumn}1`).merge();
  sheet.getRange("A1").values = [[`Aident Clinic — ${spec.name}`]];
  sheet.getRange(`A1:${lastColumn}1`).format = titleFormat;
  sheet.getRange(`A1:${lastColumn}1`).format.rowHeight = 28;
  sheet.getRange(`A3:${lastColumn}3`).values = [spec.columns];
  sheet.getRange(`A3:${lastColumn}3`).format = headerFormat;
  sheet.getRange(`A3:${lastColumn}3`).format.rowHeight = 32;
  sheet.getRange(`A4:${lastColumn}${spec.rows.length + 3}`).values = spec.rows;
  const used = sheet.getRange(`A3:${lastColumn}${spec.rows.length + 3}`);
  used.format.wrapText = true;
  used.format.verticalAlignment = "top";
  used.format.borders = { preset: "outside", style: "thin", color: "#B7D9D4" };
  sheet.getRange(`A4:${lastColumn}${spec.rows.length + 3}`).format.borders = { preset: "inside", style: "thin", color: "#E2EEEC" };
  sheet.getRange(`A4:${lastColumn}${spec.rows.length + 3}`).format.rowHeight = 42;
  if (spec.name === "Website Verification") {
    sheet.getRange("A4:I4").format.rowHeight = 90;
    sheet.getRange("A5:I5").format.rowHeight = 150;
    sheet.getRange("A6:I7").format.rowHeight = 100;
  }
  sheet.getRange(`A1:${lastColumn}${spec.rows.length + 3}`).format.autofitColumns();
  for (let col = 0; col < cols; col += 1) {
    const range = sheet.getRangeByIndexes(0, col, spec.rows.length + 3, 1);
    if (range.format.columnWidth > 34) range.format.columnWidth = 34;
    if (range.format.columnWidth < 14) range.format.columnWidth = 14;
  }
  sheet.freezePanes.freezeRows(3);
}

const verification = workbook.worksheets.getItem("Website Verification");
verification.getRange("C4:I4").format.fill = "#E6F4EA";
verification.getRange("C5:I7").format.fill = "#E6F4EA";

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/Aident-verification-report.xlsx`);

for (const spec of sheets) {
  const preview = await workbook.render({ sheetName: spec.name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/previews-${spec.name.replaceAll(" ", "-")}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const check = await workbook.inspect({ kind: "workbook,sheet,table", maxChars: 6000, tableMaxRows: 5, tableMaxCols: 10 });
console.log(check.ndjson);
