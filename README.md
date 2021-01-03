# SberHTMLParser
Converting Sberbank HTML brokerage report to Excel

<br>Using: SberHTMLParser &#60;path to file>
<br><br>.NET 5.1

<br>На текущий момент мне известно про следующие таблички в отчете
<br>
<br>"Оценка активов"
<br>"Портфель Ценных Бумаг"
<br>"Денежные средства"
<br>"Движение денежных средств за период"
<br>"Сделки купли/продажи ценных бумаг"
<br>"Движение ЦБ, не связанное с исполнением сделок"
<br>"Сделки РЕПО"
<br>"Выплаты дохода от эмитента на внешний счет"
<br>"Справочник Ценных Бумаг"

<br><br>Определение самих таблиц и параметров их идентификации производится через <a href="https://github.com/PeterKirillow/SberHTMLParser/blob/master/App.config">конфигурационный файл</a>

<br>Таким образом добавление парсинга новой таблички можно произвести путем добавления информации о ней в конфигурационный файл.
<br>Либо, при изменении формата брокерского отчета, изменить параметры идентификации таблицы

<br>После парсинга данных из отчета происходит формирование позиции на каждый день за весь период отчета в закладку "Позиция".

<br>Инсталяция происходит путем копирования папки _release на локальный диск. Может потребоваться установка .NET 5.1.
<br><a href="https://downgit.github.io/#/home?url=https://github.com/PeterKirillow/SberHTMLParser/tree/master/_release/net5.0">Или просто скачать архив</a>
