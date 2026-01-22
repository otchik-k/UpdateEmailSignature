<h1 align="center">UpdateEmailSignature</h1>
<h3>Скрипт формирует подпись для почтового клиента RoundCube по шаблону</h3>

<h4>Написан для стандартной БД RoundCube</h4>
<p>Скрипт берет из AD список sAMAccountName пользователей, получает по этим логинам значения следующих атрибутов:</p>

<p style="padding-left: 30px;">
	- cn;<br />- title;<br />
	- company;<br />
	- streetAddress;<br />
	- l;<br />
	- mail;<br />
	- telephoneNumber;<br />
	- mobile;<br />
	- pager.</p>
	
<p><br />Если атрибуты company, streetAddress, l, telephoneNumber пустые, то перезаписываются значениями, указанными в Config.txt</p>
<p>Отфильтровывает записи с пустым email и pager равным значению, которое указывается в Config.txt.</p>
<p>Выносит в отдельный список пользователей с одинаковым mail.</p>
<p>Из БД RoundCubeMail из таблицы identities достает список значений столбца email, и сравнивает со списком почтовых ящиков, полученных из AD.&nbsp;</p>
<p>Формирует список mail, котоые есть в identities, но нет в AD.</p>
<p>Формирует список mail, которые есть и в identities и в AD, но у них не совпадают cn и name (столбец в identities).</p>
<p>Формирует текст подписи для пользователей у которых cn и mail из AD совпадает с name и email из identities.</p>
<p>Если сформированный текст подписи не совпадает с тем, что храниться в identities, то скрипт перезаписывает его.</p>
<p>Так же скрипт готовит два варианта шаблона, в зависимости от наличия или отсутствия у пользователя значения в атрибуте mobile.&nbsp;</p>
<p>Все действия записываются в log\log - + "dd.MM.yyyy HH.mm.ss" + .txt.</p>