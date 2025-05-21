<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
    <!-- Устанавливаем вывод в текстовом формате с кодировкой UTF-8 -->
    <xsl:output method="text" encoding="UTF-8"/>

    <!-- Переменная, задающая имя листа данных -->
    <xsl:variable name="sheet-name" select="'Отчет'"/>

    <!-- Основной шаблон, применяемый к корневому элементу XML -->
    <xsl:template match="/">
        <!-- Заголовок отчета -->
        <xsl:text>Header_1  &#9;Date     Time      &#9;Method_Name (Filter_Name)     &#9;Calc_Mode   &#9;Sample Name                   &#9;ID Lims                    &#9;                              &#9;                              </xsl:text>

        <!-- Динамически добавляем названия элементов в заголовок -->
        <xsl:for-each select="analysis/columns/sheet[@name=$sheet-name]/column[@type='commonLine']">
            <xsl:text>&#9; &#9;</xsl:text>
            <xsl:value-of select="element"/>
            <xsl:text>&#32;&#32;&#32;&#32;&#32;&#32;&#32;</xsl:text>
        </xsl:for-each>

        <xsl:text>&#10;</xsl:text> <!-- Переход на новую строку -->

        <!-- Проходим по каждой пробе, которая видима (атрибут visible = 'yes') -->
        <xsl:for-each select="analysis/probes/probe[@visible = 'yes']">
            <!-- Сохраняем ID текущей пробы -->
            <xsl:variable name="probe_id" select="@id"/>
            <xsl:variable name="probe_name" select="@name"/>
            <xsl:variable name="name_method" select="../../titul/aname"/>

            <!-- Переменная для хранения строкового времени последней пробы -->
            <xsl:variable name="last_time">
                <xsl:for-each select="spe">
                    <xsl:variable name="timeStr" select="info/spe/timeStr"/>
                    <xsl:choose>
                        <xsl:when test="position() = last()"> <!-- Сохраняем только последнее время -->
                            <xsl:value-of select="$timeStr"/>
                        </xsl:when>
                    </xsl:choose>
                </xsl:for-each>
            </xsl:variable>

            <!-- Переменная для хранения времени последней пробы чтобы заполнить ID Lims -->
            <xsl:variable name="last_time_idlims">
                <xsl:for-each select="spe">
                    <xsl:variable name="time" select="info/spe/time"/>
                    <xsl:choose>
                        <xsl:when test="position() = last()"> <!-- Сохраняем только последнее время -->
                            <xsl:value-of select="$time"/>
                        </xsl:when>
                    </xsl:choose>
                </xsl:for-each>
            </xsl:variable>

            <!-- Сохраняем количество элементов <spe> -->
            <xsl:variable name="spe_count" select="count(spe)"/>

            <!-- Обрабатываем данные для каждой параллели <spe> -->
            <xsl:for-each select="spe">
                <xsl:variable name="cl_id" select="@id"/>
                <xsl:variable name="timeStr" select="info/spe/timeStr"/>

                <!-- Строка данных для измерения Single -->
                <xsl:text>Single    &#9;</xsl:text>
                <xsl:value-of select="concat($timeStr, substring('                              ', 1, 19 - string-length($timeStr)))"/>
                <xsl:text>&#9;</xsl:text>
                <xsl:value-of select="concat($name_method, substring('                              ', 1, 30 - string-length($name_method)))"/>
                <xsl:text>&#9;AutoResult  &#9;</xsl:text>
                <xsl:value-of select="concat($probe_name, substring('                              ', 1, 30 - string-length($probe_name)))"/>
                <xsl:text>&#9;</xsl:text>
                <xsl:value-of select="concat($last_time_idlims, substring('                              ', 1, 30 - string-length($last_time_idlims)))"/>
                <xsl:text>&#9;                              &#9;                              </xsl:text>

                <!-- Проходим по каждому элементу, чтобы вывести значения -->
                <xsl:for-each select="../../../columns/sheet[@name=$sheet-name]/column[@type='commonLine']">
                    <!-- Проверка наличия значения для текущей параллели -->
                    <xsl:choose>
                        <xsl:when test="cells/pc/cl[@i=$cl_id]">
                            <xsl:variable name="value" select="cells/pc/cl[@i=$cl_id]/@v"/>
                            <xsl:text>&#9;</xsl:text>
                            <xsl:choose>
                                <!-- Проверка наличия знака '<' -->
                                <xsl:when test="starts-with($value, '&lt;')">
                                    <xsl:text>&lt;&#9;</xsl:text>
                                    <xsl:variable name="value_clear" select="substring-after($value, '&lt;')"/>
                                    <xsl:call-template name="format-value">
                                        <xsl:with-param name="value" select="$value_clear"/>
                                    </xsl:call-template>
                                </xsl:when>
                                <!-- Проверка наличия знака '>' -->
                                <xsl:when test="starts-with($value, '&gt;')">
                                    <xsl:text>&gt;&#9;</xsl:text>
                                    <xsl:variable name="value_clear" select="substring-after($value, '&gt;')"/>
                                    <xsl:call-template name="format-value">
                                        <xsl:with-param name="value" select="$value_clear"/>
                                    </xsl:call-template>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:text> &#9;</xsl:text>
                                    <xsl:variable name="value_clear" select="$value"/>
                                    <xsl:call-template name="format-value">
                                        <xsl:with-param name="value" select="$value_clear"/>
                                    </xsl:call-template>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:text>         &#9; &#9;</xsl:text>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:for-each>
                <xsl:text>&#10;</xsl:text>
            </xsl:for-each>

            <!-- Проверяем, есть ли элементы <spe>, прежде чем выводить Average -->
            <xsl:if test="$spe_count > 0">
                <!-- Выводим строку данных пробы -->
                <xsl:text>Average   &#9;</xsl:text>
                <xsl:value-of select="concat($last_time, substring('                              ', 1, 19 - string-length($last_time)))"/>
                <xsl:text>&#9;</xsl:text>
                <xsl:value-of select="concat($name_method, substring('                              ', 1, 30 - string-length($name_method)))"/>
                <xsl:text>&#9;AutoResult  &#9;</xsl:text>
                <xsl:value-of select="concat($probe_name, substring('                              ', 1, 30 - string-length($probe_name)))"/>
                <xsl:text>&#9;</xsl:text>
                <xsl:value-of select="concat($last_time_idlims, substring('                              ', 1, 30 - string-length($last_time_idlims)))"/>
                <xsl:text>&#9;                              &#9;                              </xsl:text>

                <!-- Проходим по каждому элементу с типом commonLine -->
                <xsl:for-each select="../../columns/sheet[@name=$sheet-name]/column[@type='commonLine']">
                    <!-- Проверяем наличие данных в <pc> с соответствующим ID -->
                    <xsl:choose>
                        <xsl:when test="cells/pc[@i=$probe_id]">
                            <!-- Если элемент найден, форматируем значение -->
                            <xsl:variable name="value" select="cells/pc[@i=$probe_id]/@v"/>
                            <xsl:text>&#9;</xsl:text>
                            <xsl:choose>
                                <!-- Проверка наличия знака '<' -->
                                <xsl:when test="starts-with($value, '&lt;')">
                                    <xsl:text>&lt;&#9;</xsl:text>
                                    <xsl:variable name="value_clear" select="substring-after($value, '&lt;')"/>
                                    <xsl:call-template name="format-value">
                                        <xsl:with-param name="value" select="$value_clear"/>
                                    </xsl:call-template>
                                </xsl:when>
                                <!-- Проверка наличия знака '>' -->
                                <xsl:when test="starts-with($value, '&gt;')">
                                    <xsl:text>&gt;&#9;</xsl:text>
                                    <xsl:variable name="value_clear" select="substring-after($value, '&gt;')"/>
                                    <xsl:call-template name="format-value">
                                        <xsl:with-param name="value" select="$value_clear"/>
                                    </xsl:call-template>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:text> &#9;</xsl:text>
                                    <xsl:variable name="value_clear" select="$value"/>
                                    <xsl:call-template name="format-value">
                                        <xsl:with-param name="value" select="$value_clear"/>
                                    </xsl:call-template>
                                </xsl:otherwise>
                            </xsl:choose>
                        </xsl:when>
                        <xsl:otherwise>
                            <!-- Если значение отсутствует, выводим пустое поле -->
                            <xsl:text>         &#9; &#9;</xsl:text>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:for-each>
                <!-- Переход на новую строку после каждой пробы -->
                <xsl:text>&#10;</xsl:text>
            </xsl:if>
        </xsl:for-each>
    </xsl:template>

    <!-- Шаблон для форматирования значений концентраций -->
    <xsl:template name="format-value">
        <xsl:param name="value"/>
        <xsl:variable name="number" select="translate($value, ',', '.')" /> <!-- Заменяем запятые на точки для обработки как чисел -->

        <xsl:choose>
            <!-- Если значение содержит 'e' или 'E', предполагаем научную нотацию -->
            <xsl:when test="contains($number, 'e') or contains($number, 'E')">
                <xsl:variable name="mantissa" select="number(translate(substring-before($value, 'e'),  ',', '.'))"/>
                <xsl:variable name="exponent" select="substring-after($value, 'e')"/>
                <!-- Если экспонента отрицательная -->
                <xsl:choose>
                    <xsl:when test="number($exponent) &lt; 0">
                        <!-- Количество ведущих нулей на основе значения экспоненты -->
                        <xsl:variable name="zeros" select="substring('0000000000', 1, -1 * number($exponent) - 1)"/>
                        <!-- Форматирование мантиссы -->
                         <xsl:variable name="number_exp" select="concat('0.', $zeros, number(translate($mantissa,'.', '')))" />

                        <xsl:value-of select="translate(format-number(number($number_exp), '00.000000'), '.', ',')"/>
                    </xsl:when>
                    <xsl:when test="number($exponent) = 0">
                        <xsl:value-of select="translate(format-number($mantissa, '00.000000'), '.', ',')"/>
                    </xsl:when>

                    <!-- Обработка положительной экспоненты (вывод мантиссы напрямую) -->
                    <xsl:otherwise>
                        <!-- Количество ведущих нулей на основе значения экспоненты -->
                        <xsl:variable name="zeros" select="substring('0000000000', 1, number(substring-after($exponent, '+')))"/>
                        <xsl:variable name="factor" select="number(concat('1', $zeros))"/>
                        <xsl:value-of select="translate(format-number(($mantissa * $factor), '00.000000'), '.', ',')"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$number != '' and number($number) = $number"> <!-- Проверяем, является ли значение числом -->
                <xsl:value-of select="translate(format-number(number($number), '00.000000'),  '.', ',')"/> <!-- Форматируем число -->
            </xsl:when>
            <xsl:otherwise>
                <xsl:text>&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;&#32;</xsl:text> <!-- Если значение пустое или не число, выводим пустое поле -->
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

</xsl:stylesheet>
