# metal-calculator
## Простая программка для парсинга веса латунных или алюминиевых листов с сайта metal-calculator.ru

Программа для решения такой задачи:

Нужно пройтись по товарам и проставить характеристику "вес", где она не указана.
Вес берем с калькулятора металлопроката: https://metal-calculator.ru/
Там выбираем "Алюминий", "Лист/плита", указываем марку (если такой марки там нет - то Прочее), толщину, ширину, длину, количество (1 шт) и он показывает вес.

Пример на скриншоте:
screen.jpg

Забираем вес и забиваем данными в табличку Excel, как в примере:
tz.csv

Программа должна работать с латунными или алюминиевыми листами, имеющимися в калькуляторе, и выдавать вес с учетом характеристик. Данные в файле просто для примера.

В принципе, программу можно доработать под любой металл и вид проката со своим набором характеристик.
 
Есть возможность получения на выходе файла .exe с симпатичным интерфейсом (но мне это не нужно на данный момент).
