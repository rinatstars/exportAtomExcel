<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

    <!-- Определяем формат выходного XML: кодировка UTF-8, с отступами -->
    <xsl:output method="xml" encoding="UTF-8" indent="yes"/>

    <!-- Переменная, которая задает имя листа с данными. Это значение можно изменить для работы с другим листом -->
    <xsl:variable name="sheet-name" select="'Отчет'"/> <!-- МЕСТО ДЛЯ НАСТРОЙКИ -->

    <!-- Основной шаблон, который применяется к корневому элементу XML -->
    <xsl:template match="/">
        <results>
            <!-- Создаем заголовок с информацией о заказчике, исполнителе, дате и методе анализа -->

            <header>
                <!-- Блок с именем заказчика -->
                <customer>
                    <!-- Вытягиваем информацию о заказчике из поля "comment" в XML (обычно содержит текст "Заказчик: ...") -->
                    <xsl:value-of select="analysis/titul/comment"/>
                </customer>

                <!-- Блок с именем исполнителя анализа -->
                <executor>
                    <!-- Вытягиваем имя исполнителя из поля "user" -->
                    <xsl:value-of select="analysis/titul/user"/>
                </executor>

                <!-- Блок с датой проведения анализа -->
                <date>
                    <!-- Вытягиваем дату из поля "date_str" -->
                    <xsl:value-of select="analysis/titul/date_str"/>
                </date>

                <!-- Блок с названием метода анализа -->
                <method>
                    <!-- Вытягиваем название метода из поля "aname" -->
                    <xsl:value-of select="analysis/titul/aname"/>
                </method>

                <!-- Блок с наименованием прибора -->
                <device>
                    <!-- Вытягиваем название метода из поля "aname" -->
                    <xsl:value-of select="analysis/titul/device"/>
                </device>

                <!-- Блок с названием организации -->
                <organization>
                    <!-- Вытягиваем название метода из поля "aname" -->
                    <xsl:value-of select="analysis/titul/organization"/>
                </organization>
            </header>

            <!-- Обрабатываем все пробы, которые находятся в элементе "probes" -->
            <xsl:for-each select="analysis/probes/probe[@visible = 'yes']">
                <!-- Сохраняем ID текущей пробы в переменную -->
                <xsl:variable name="probe_id" select="@id"/>

                <probe>
                    <!-- Вытягиваем имя пробы из атрибута "name" -->
                    <name>
                        <xsl:value-of select="@name"/>
                    </name>

                    <!-- Проходим по всем колонкам с типом "commonLine" -->
                    <xsl:for-each select="../../columns/sheet[@name=$sheet-name]/column[@type='commonLine']">
                        <element>
                            <name>
                                <!-- Название элемента из тега <element> -->
                                <xsl:value-of select="element"/>
                            </name>

                            <!-- Проверяем наличие <pc> с соответствующим ID -->
                            <xsl:choose>
                                <xsl:when test="cells/pc[@i=$probe_id]">
                                    <!-- Если элемент <pc> найден, выводим значение @v -->
                                    <!-- Среднее -->
                                    <mean>
                                        <xsl:value-of select="cells/pc[@i=$probe_id]/@v"/>
                                    </mean>

                                    <!-- Параллельные измерения -->
                                    <parallel>
                                        <xsl:for-each select="cells/pc[@i=$probe_id]/cl">
                                            <measurement>
                                                <xsl:value-of select="@v"/>
                                            </measurement>
                                        </xsl:for-each>
                                    </parallel>
                                </xsl:when>
                                <xsl:otherwise>
                                    <!-- Если элемент <pc> не найден, добавляем пустое значение -->
                                    <value></value>
                                </xsl:otherwise>
                            </xsl:choose>
                        </element>
                    </xsl:for-each>
                </probe>
            </xsl:for-each>
        </results>
    </xsl:template>
</xsl:stylesheet>
