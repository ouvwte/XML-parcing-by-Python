PROCEDURE PRO_UNLOAD_XML_RLW is 
    l_peremen    CLOB; --:= EMPTY_CLOB (); -- файл шаблона
    l_file       BLOB; --:= EMPTY_BLOB (); -- итоговый файл
    l_row_d      CLOB; --:= EMPTY_CLOB (); -- блоки информации для добавления в файл шаблона
/*    l_peremen_2    CLOB;
    l_peremen_3    CLOB;
    l_peremen_4    CLOB;*/
    v_cou NUMBER;
    v_row_end NUMBER;
    v_sum_pack NUMBER;
    v_summa_ac NUMBER;
    v_sum_quan_fakt NUMBER;
    v_sum_weight_net NUMBER;
    v_sum_weight_gross NUMBER;
    naim_rus varchar2(2000);
    naim_eng varchar2(2000);
    v_str varchar2(2000);
    v_str_date varchar2(2000);
    v_id_transaction number(10);
--    v_row_end_1 number;
--    v_row_end_2 number;
    v_name_invoice varchar2(2000);
    v_weight_nett NUMBER;
    v_weight_gross NUMBER;


/***********************************************************************************************************
    В ДАННОЙ ПРОЦЕДУРЕ ПРОИСХОДИТ ВЫГРУЗКА НОВОЙ ГОДОВОЙ ДОСКИ ПРОИЗВОДСТВЕННОГО АНАЛИЗА ДЛЯ RAILWAY
************************************************************************************************************/

--КУРСОР, КОТОРЫЙ ВЫТАСКИВАЕТ ДАННЫЕ ДЛЯ ОТОБРАЖЕНИЯ
CURSOR Q_1 IS
    SELECT * from railway.rlw_board
        where god=2025 and mes=9;
    
        
--НАЧАЛО ПРОЦЕДУРЫ
BEGIN

    BEGIN 
SELECT t.file_excel INTO l_peremen
      FROM product.pro_file_excel t
     WHERE t.name_excel = 'unload_xml_rlw'
       AND t.num_excel = 8;
  EXCEPTION WHEN OTHERS THEN
       l_peremen:=NULL;
 END;
       
    -- очищаем BLOB Excel файла
     dbms_output.put_line('+++ '/*|| l_peremen*/);
BEGIN     
    UPDATE product.pro_file_excel t SET t.blob_excel = ''
     WHERE t.name_excel = 'unload_xml_rlw'
       AND t.num_excel = 8;
    COMMIT;
  --  EXCEPTION WHEN OTHERS THEN
       NULL;
 END;
  dbms_output.put_line('*** '/*|| l_peremen*/);

    dbms_output.put_line('/// '/*|| l_peremen*/);
execute immediate 'ALTER SESSION SET NLS_NUMERIC_CHARACTERS = ''. ''';   
 
 -- НАЧАЛО ФОРМИРОВАНИЯ КОДА XML-ФАЙЛА    
    l_row_d:= '<?xml version="1.0" encoding="UTF-8"?>
   <?mso-application progid="Excel.Sheet"?>
   <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
   xmlns:o="urn:schemas-microsoft-com:office:office"
   xmlns:x="urn:schemas-microsoft-com:office:excel"
   xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
   xmlns:html="http://www.w3.org/TR/REC-html40">
  <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
   <Created>2015-06-05T18:19:34Z</Created>
   <LastSaved>2015-06-05T18:19:39Z</LastSaved>
   <Version>16.00</Version>
  </DocumentProperties>
 
  <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
   <AllowPNG/>
 </OfficeDocumentSettings>
 
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>12300</WindowHeight>
  <WindowWidth>28800</WindowWidth>
  <WindowTopX>0</WindowTopX>
  <WindowTopY>0</WindowTopY>
  <RefModeR1C1/>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook> ';
 l_peremen:=CONVERT(l_row_d, 'CL8ISO8859P5');
 dbms_output.put_line('111 '/*|| l_peremen*/);

-- СТИЛИ 
 l_row_d := '
 <Styles> <!-- СТИЛИ -->
  <Style ss:ID="m461347280">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="11"
      ss:Color="#000000" ss:Bold="1"/>
     <Interior ss:Color="#F8CBAD" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="m461347300">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="12"
      ss:Color="#000000"/>
     <Interior ss:Color="#F8CBAD" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="m461349208">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="12"
      ss:Color="#000000" ss:Bold="1"/>
     <Interior ss:Color="#F8CBAD" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s39">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="12"
      ss:Color="#000000" ss:Bold="1"/>
     <Interior ss:Color="#B4C6E7" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s47">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="12"
      ss:Color="#000000"/>
     <Interior ss:Color="#C6E0B4" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s48">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="12"
      ss:Color="#000000"/>
     <Interior ss:Color="#FFE699" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s64">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="12"
      ss:Color="#000000" ss:Italic="1"/>
     <Interior ss:Color="#B4C6E7" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s70">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="12"
      ss:Color="#000000"/>
     <Interior ss:Color="#F8CBAD" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s71">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
     <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="12"
      ss:Color="#000000"/>
     <Interior ss:Color="#F8CBAD" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s79">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
      ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="11"
      ss:Color="#000000" ss:Bold="1"/>
     <Interior ss:Color="#C6E0B4" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s80">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
      ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="11"
      ss:Color="#000000" ss:Bold="1"/>
     <Interior ss:Color="#FFE699" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s83">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
      ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="11"
      ss:Color="#000000" ss:Bold="1"/>
     <Interior ss:Color="#B4C6E7" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s84">
     <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90"
      ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="11"
      ss:Color="#000000" ss:Bold="1"/>
     <Interior ss:Color="#B4C6E7" ss:Pattern="Solid"/>
    </Style>
  <Style ss:ID="s152">
     <Alignment ss:Horizontal="Right" ss:Vertical="Center" ss:WrapText="1"/>
     <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
     </Borders>
     <Font ss:FontName="Calibri" x:CharSet="204" x:Family="Swiss" ss:Size="11"
      ss:Color="#000000" ss:Bold="1"/>
     <Interior ss:Color="#F8CBAD" ss:Pattern="Solid"/>
    </Style>
 </Styles>';
 DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));
 dbms_output.put_line('2222222222222 ');
 
  
-- НАЧИНАЕМ ФОРМИРОВАТЬ ОСНОВНЫЕ ВЫХОДНЫЕ ДАННЫЕ 
-- ЛИСТ "ФЕВРАЛЬ" - КАК ПРИМЕР (ПОКА ОДИН)
  l_row_d := ' 
 <Worksheet ss:Name="февраль">
  <Table ss:ExpandedColumnCount="93" ss:ExpandedRowCount="37" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:AutoFitWidth="0" ss:Width="81"/>
   <Column ss:AutoFitWidth="0" ss:Width="191.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="39"/>
   <Column ss:AutoFitWidth="0" ss:Width="43.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="37.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="35.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="34.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="33.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="36"/>
   <Column ss:AutoFitWidth="0" ss:Width="34.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="33.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="33"/>
   <Column ss:AutoFitWidth="0" ss:Width="37.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="29.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="30"/>
   <Column ss:AutoFitWidth="0" ss:Width="33.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="37.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="33"/>
   <Column ss:AutoFitWidth="0" ss:Width="30"/>
   <Column ss:AutoFitWidth="0" ss:Width="33"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="33.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="33"/>
   <Column ss:AutoFitWidth="0" ss:Width="34.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="32.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="33.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="35.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="36"/>
   <Column ss:AutoFitWidth="0" ss:Width="33"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="40.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="41.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="39.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="36.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="41.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="44.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="36"/>
   <Column ss:AutoFitWidth="0" ss:Width="39"/>
   <Column ss:AutoFitWidth="0" ss:Width="42"/>
   <Column ss:AutoFitWidth="0" ss:Width="47.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="39.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="42"/>
   <Column ss:AutoFitWidth="0" ss:Width="43.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="41.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="43.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="41.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="45.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="39.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="39.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="40.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="36.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="39.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="43.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="40.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="39.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="43.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="40.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="41.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="40.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="34.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="39.75"/>
   <Column ss:Index="69" ss:AutoFitWidth="0" ss:Width="42"/>
   <Column ss:AutoFitWidth="0" ss:Width="34.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="36.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="40.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="44.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="39.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="40.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="43.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="42"/>
   <Column ss:AutoFitWidth="0" ss:Width="41.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="44.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="37.5"/>
   <Column ss:Hidden="1" ss:AutoFitWidth="0" ss:Span="7"/>
   <Row ss:Height="15.75">
    <Cell ss:MergeAcross="1" ss:StyleID="s152"><Data ss:Type="String">Число</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349208"><Data ss:Type="Number">3</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349228"><Data ss:Type="Number">4</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349248"><Data ss:Type="Number">5</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349268"><Data ss:Type="Number">6</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349288"><Data ss:Type="Number">7</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349308"><Data ss:Type="Number">10</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349328"><Data ss:Type="Number">11</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349348"><Data ss:Type="Number">12</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349368"><Data ss:Type="Number">13</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349388"><Data ss:Type="Number">14</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349168"><Data ss:Type="Number">17</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349516"><Data ss:Type="Number">18</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349536"><Data ss:Type="Number">19</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349556"><Data ss:Type="Number">20</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349576"><Data ss:Type="Number">21</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349596"><Data ss:Type="Number">24</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349616"><Data ss:Type="Number">25</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349636"><Data ss:Type="Number">26</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349656"><Data ss:Type="Number">27</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349676"><Data ss:Type="Number">28</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349696"><Data ss:Type="Number">29</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="m461349716"><Data ss:Type="Number">30</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m461349496"><Data ss:Type="String">Итого</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="30">
    <Cell ss:MergeAcross="1" ss:StyleID="m461347280"><Data ss:Type="String">Клиенты</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s83"><Data ss:Type="String">откл-е</Data></Cell>
    <Cell ss:StyleID="s84"><Data ss:Type="String">*</Data></Cell>
    <Cell ss:StyleID="s79"><Data ss:Type="String">план</Data></Cell>
    <Cell ss:StyleID="s80"><Data ss:Type="String">факт</Data></Cell>
    <Cell ss:StyleID="s81"><Data ss:Type="String">откл-е</Data></Cell>
   </Row>
   '; 

 DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5')); 
 dbms_output.put_line('3333333333333 ');

 for rec in Q_1 loop
 l_row_d := '
   <Row ss:AutoFitHeight="0" ss:Height="16.5">
    <Cell ss:MergeDown="1" ss:StyleID="s21"><Data ss:Type="String">Item</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s21"><Data ss:Type="String">HS Code</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s21"><Data ss:Type="String">Part reference №</Data></Cell>
    <Cell ss:MergeAcross="3" ss:MergeDown="1" ss:StyleID="s21"><Data
      ss:Type="String">Part name</Data></Cell>
    <Cell ss:StyleID="s21"><Data ss:Type="String">Quantity</Data></Cell>
    <Cell ss:StyleID="s21"><Data ss:Type="String">Price CIF Danang</Data></Cell>
    <Cell ss:StyleID="s21"><Data ss:Type="String">Contract         amount</Data></Cell>
    <Cell ss:StyleID="s21"><Data ss:Type="String">Weight   &#10;Net                                            </Data></Cell>
    <Cell ss:StyleID="s21"><Data ss:Type="String">Weight                                             Gross</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s21"><Data ss:Type="String">Country of origin</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s21"><Data ss:Type="String">Package of №</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s21"><Data ss:Type="String">Annex</Data></Cell>
   </Row>
   <Row ss:Height="16.5">
    <Cell ss:Index="8" ss:StyleID="s211"><Data ss:Type="String">Кол-во</Data></Cell>
    <Cell ss:StyleID="s211"><Data ss:Type="String">Цена контр.</Data></Cell>
    <Cell ss:StyleID="s211"><Data ss:Type="String">Сумма контр.</Data></Cell>
    <Cell ss:StyleID="s211"><Data ss:Type="String">Вес нетто</Data></Cell>
    <Cell ss:StyleID="s211"><Data ss:Type="String">Вес брутто</Data></Cell>
   </Row>
   <Row ss:Height="24.75">
    <Cell ss:StyleID="s13"><Data ss:Type="String">Позиция</Data></Cell>
    <Cell ss:StyleID="s13"/>
    <Cell ss:StyleID="s13"><Data ss:Type="String">Номер детали</Data></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="s13"><Data ss:Type="String">Наименование детали</Data></Cell>
    <Cell ss:StyleID="s13"><Data ss:Type="String">unit/шт</Data></Cell>
    <Cell ss:StyleID="s13"><Data ss:Type="String">USD</Data></Cell>
    <Cell ss:StyleID="s13"><Data ss:Type="String">USD</Data></Cell>
    <Cell ss:StyleID="s13"><Data ss:Type="String">kg/кг</Data></Cell>
    <Cell ss:StyleID="s13"><Data ss:Type="String">kg</Data></Cell>
    <Cell ss:StyleID="s13"><Data ss:Type="String">Страна происхождения</Data></Cell>
    <Cell ss:StyleID="s13"><Data ss:Type="String">№ места</Data></Cell>
    <Cell ss:StyleID="s13"><Data ss:Type="String">№ приложения</Data></Cell>
   </Row>';
   DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));
   dbms_output.put_line('4444444444 ');
  end loop;
   
for rec in Q00 LOOP   
    l_row_d := '   
   <Row ss:Height="20.5">
    <Cell ss:StyleID="s12"><Data ss:Type="Number">' || Q00%ROWCOUNT || '</Data></Cell>
    <Cell ss:StyleID="s12"><Data ss:Type="String"></Data></Cell>
    <Cell ss:StyleID="s12"><Data ss:Type="String">' || rec.num_obj || '</Data></Cell>      
    <Cell ss:MergeAcross="3" ss:StyleID="s12"><Data ss:Type="String">' || rec.name_obj || '</Data></Cell>
    <Cell ss:StyleID="s12"><Data ss:Type="Number">' || rec.quan_det || '</Data></Cell>
    <Cell ss:StyleID="s12"><Data ss:Type="Number">' || NVL(rec.price_ak_val, rec.price) || '</Data></Cell>
    <Cell ss:StyleID="s12"><Data ss:Type="Number">' || round(rec.summa_ac,2) || '</Data></Cell>
    <Cell ss:StyleID="s12"><Data ss:Type="Number">' || round(rec.weight_nett,3) || '</Data></Cell>
    <Cell ss:StyleID="s12"><Data ss:Type="Number">' || round(rec.weight_gross,3) || '</Data></Cell>
    <Cell ss:StyleID="s12"><Data ss:Type="String">' || rec.name_country_out || '</Data></Cell>
    <Cell ss:StyleID="s12"><Data ss:Type="Number"></Data></Cell>     <!-- Package of № -->
    <Cell ss:StyleID="s12"><Data ss:Type="Number"></Data></Cell>      <!-- Annex -->
    <Cell ss:Index="17" ss:StyleID="s23"/>
   </Row>';
    DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));
end loop;
    dbms_output.put_line('end loop spec 555555555');

/*begin
    select distinct a.id_transaction
    into v_id_transaction
    FROM pro_invoice a,
         agr_contractor c,
         agr_contractor d,
         agr_agreement ag,
         agr_city ct,
          rus_country cntr,
          pla_ak_appl  ak,
          pro_inv_ak_appl ap
   WHERE -- a.id_dept_owner = 113396 and
    ap.id_invoice=a.id_invoice
     AND a.id_invoice = p_id_invoice                                     --38467413
     AND pro_rep.ctr_by_url (a.id_dept_owner) = c.id_contr
     AND a.id_contr_in = d.id_contr
     AND a.id_agr = ag.id_agr
  ---   AND a.id_doc=4301
     AND ct.code=c.code_city
     AND cntr.code=ct.code_country
     AND ak.id_pla_ak_appl=ap.id_pla_ak_appl; 
--end;*/
-- dbms_output.put_line(v_id_transaction);

 
 for rec in Q00 loop
    begin
        v_sum_quan_fakt:=nvl(v_sum_quan_fakt,0)+nvl(rec.quantity_fakt,0);
        v_sum_weight_net:=nvl(v_sum_weight_net,0)+nvl(rec.weight_nett,0);
        v_sum_weight_gross:=nvl(v_sum_weight_gross,0)+nvl(rec.weight_gross,0);
        v_summa_ac:=nvl(v_summa_ac,0)+nvl(rec.summa_ac,0);
        v_row_end:=Q00%ROWCOUNT+26;
    end;
end loop; 

    begin
        select count(id_doc_stock) into v_cou
        from (select distinct a.id_doc_stock
          from stc_doc_stock a, pla_ak_appl b, pro_inv_ak_appl d
         where a.id_source=b.id_pla_ak_appl
           and a.data_source=-39
           and a.id_doc=3119
           and b.id_pla_ak_appl=d.id_pla_ak_appl
           and d.id_invoice=p_id_invoice);
    exception when others then
	   v_cou:=0;
    end;
 
 --dbms_output.put_line('end loop invoice');
 
---naim_rus:= ' ООО "Автомобильный завод ГАЗ" , Россия '; -- убрала Баженова, присвоено в курсоре Q_zag
---naim_eng:= 'GAZ Automobile plant LLC, Russia';

 l_row_d := '
   <Row ss:Height="14.25">
    <Cell ss:MergeAcross="7" ss:StyleID="s13"><ss:Data ss:Type="String">TOTAL</ss:Data></Cell>
    <Cell ss:StyleID="s13"/>
    <Cell ss:StyleID="s13"><Data ss:Type="Number">' || round(v_summa_ac,2) || '</Data></Cell>
    <Cell ss:StyleID="s13"><Data ss:Type="Number">' || round(v_sum_weight_net,3) || '</Data></Cell>
    <Cell ss:StyleID="s13"><Data ss:Type="Number">' || round(v_sum_weight_gross,3) || '</Data></Cell>
    <Cell ss:StyleID="s13"/>
    <Cell ss:StyleID="s13"/>
    <Cell ss:StyleID="s13"/>
    <Cell ss:Index="17" ss:StyleID="s23"/>
   </Row>
   <Row ss:AutoFitHeight="0"></Row>';
   DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));
  dbms_output.put_line('end sheet spec 6666666666 ');
   
 --БЕСПЛАТНЫЙ ДОСЫЛ
if v_id_transaction<>91764 then
    --begin
    l_row_d := '
        <Row ss:AutoFitHeight="0" ss:Height="10.5"></Row> 
        <Row ss:AutoFitHeight="0" ss:Height="10.5">
         <Cell ss:MergeAcross="8" ss:StyleID="s3"><Data ss:Type="String">Примечание: Поставка на безвозмездной основе, цены указаны для проведения таможенного оформления</Data></Cell>
        </Row>
        <Row ss:AutoFitHeight="0" ss:Height="10.5">
         <Cell ss:MergeAcross="8" ss:StyleID="s3"><Data ss:Type="String">Note: the Supply on a free of charge basis, the prices are specified for customs registration</Data></Cell>
        </Row>';
        
DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));
dbms_output.put_line('777777777777');
end if; 
   
 l_row_d := '
   <Row ss:AutoFitHeight="0"></Row>
   <Row ss:AutoFitHeight="0"></Row>

   <Row ss:Height="14.25">
    <Cell ss:StyleID="s9"><Data ss:Type="String">Quantity of tare:</Data></Cell>
    <Cell ss:StyleID="s9"/>
    <Cell ss:Index="7" ss:MergeAcross="2" ss:MergeDown="1" ss:StyleID="s7"><Data ss:Type="Number">' || v_cou || '</Data></Cell>
    <Cell ss:StyleID="s23"/>
    <Cell ss:StyleID="s23"/>
    <Cell ss:StyleID="s23"/>
    <Cell ss:StyleID="s23"/>
    <Cell ss:StyleID="s23"/>
    <Cell ss:StyleID="s23"/>
    <Cell ss:StyleID="s23"/>
    <Cell ss:StyleID="s23"/>
    <Cell ss:StyleID="s23"/>
    <Cell ss:StyleID="s23"/>
   </Row>
   <Row>
    <Cell ss:StyleID="s55"><Data ss:Type="String">Количество тары:</Data></Cell>
   </Row>
   <Row ss:Height="14.25">
    <Cell ss:Index="11" ss:StyleID="s55"><Data ss:Type="String">Signatures:</Data></Cell>
   </Row>
   <Row>
    <Cell ss:Index="11" ss:MergeAcross="4" ss:StyleID="s10"><Data ss:Type="String">' ||  naim_rus || '</Data></Cell>
   </Row>
   <Row>
    <Cell ss:StyleID="s9"><Data ss:Type="String">Quantity of packing cases</Data></Cell>
    <Cell ss:StyleID="s9"/>
    <Cell ss:Index="7" ss:MergeAcross="2" ss:MergeDown="1" ss:StyleID="s7"><Data ss:Type="Number">' || v_cou || '</Data></Cell>
    <Cell ss:StyleID="s23"/>
    <Cell ss:MergeAcross="4" ss:StyleID="s11"><Data ss:Type="String">' || naim_eng || '</Data></Cell>
    <Cell ss:StyleID="s23"/>
   </Row>
   <Row>
    <Cell ss:StyleID="s55"><Data ss:Type="String">Количество мест</Data></Cell>
    <Cell ss:Index="11" ss:MergeAcross="4" ss:StyleID="s11"/>
   </Row>
   <Row>
    <Cell ss:Index="11" ss:MergeAcross="4" ss:StyleID="s11"/>
   </Row>
   <Row>
    <Cell ss:Index="11" ss:MergeAcross="4" ss:StyleID="s11"/>
   </Row>
   <Row>
    <Cell ss:Index="11" ss:MergeAcross="4" ss:StyleID="s11"/>
   </Row>
   
   <Row>
	<Cell ss:Index="11" ss:MergeAcross="4" ss:StyleID="s13"/>
   </Row> ';
DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));
  dbms_output.put_line('end sheet spec 777777777777 ');
 
--==============================================================================
 
 l_row_d := ' 
 </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.31496062992125984"/>
    <Footer x:Margin="0.31496062992125984"/>
    <PageMargins x:Bottom="0.74803149606299213" x:Left="0.11811023622047245"
     x:Right="0.11811023622047245" x:Top="0.74803149606299213"/>
   </PageSetup>
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Zoom>120</Zoom>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>22</ActiveRow>
     <RangeSelection>R23C1:R29C15</RangeSelection>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>';
DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));

  dbms_output.put_line('8888888888888');
  
  
 for rec in Q_zag loop
  l_row_d := ' 
  <Worksheet ss:Name="Инвойс">	<!-- ЛИСТ ИНВОЙС -->
  <Table ss:ExpandedColumnCount="9" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:AutoFitWidth="0" ss:Width="82.5"/>
   <Column ss:AutoFitWidth="0" ss:Width="129"/>
   <Column ss:Index="4" ss:AutoFitWidth="0" ss:Width="30"/>
   <Column ss:AutoFitWidth="0" ss:Width="54"/>
   <Column ss:AutoFitWidth="0" ss:Width="39.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="33"/>
   <Column ss:AutoFitWidth="0" ss:Width="27.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="59.25"/>
   <Row ss:AutoFitHeight="0" ss:Height="16.5">
    <Cell ss:StyleID="s1"><Data ss:Type="String">' || rec.eng_name_contr_out || '</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s2"><Data ss:Type="String">' || v_name_invoice || ' №</Data></Cell>
    <Cell ss:Index="7" ss:StyleID="s4"><Data ss:Type="Number">' || rec.doc_number || '</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s1"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s3"><Data ss:Type="String">' || rec.rus_name_contr_out || '</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s8"><Data ss:Type="String">СЧЕТ</Data></Cell>
    <Cell ss:StyleID="s2"><Data ss:Type="String">№</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s1"><Data ss:Type="String">' || rec.eng_address_contr_out || '</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s1"/>
    <Cell ss:Index="7" ss:StyleID="s1"><Data ss:Type="String">Date </Data></Cell>
    <Cell ss:StyleID="s4"><Data ss:Type="String">' || rec.doc_date || '</Data></Cell> 
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s3"><Data ss:Type="String">' || rec.rus_address_contr_out || '</Data></Cell>
    <Cell ss:Index="7" ss:StyleID="s5"><Data ss:Type="String">Дата</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:MergeAcross="1" ss:StyleID="s1"><Data ss:Type="String">tel / fax: </Data></Cell>
   </Row>
   <Row ss:Index="10" ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s1"><Data ss:Type="String">Consignee</Data></Cell>
    <Cell ss:MergeDown="4" ss:StyleID="s8"><Data ss:Type="String">' || rec.eng_name_contr_in || '</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s1"><Data ss:Type="String">Buyer</Data></Cell>
    <Cell ss:Index="6" ss:MergeAcross="3" ss:MergeDown="4" ss:StyleID="s8"><Data
      ss:Type="String">' || rec.eng_address_contr_in || '</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">Грузополучатель</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s5"><Data ss:Type="String">Покупатель</Data></Cell>
   </Row>
   <Row ss:Index="16" ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s9"><Data ss:Type="String">Consignor</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s8"><Data ss:Type="String">' || rec.eng_name_contr_out || '</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">Грузоотправитель</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s9"><Data ss:Type="String">Departure </Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s8"><Data ss:Type="String">' || rec.dep || '</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">Пункт отправления</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s9"><Data ss:Type="String">Destination</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s8"><Data ss:Type="String">' || rec.dest || '</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s9"><Data ss:Type="String">Contract</Data></Cell>
    <Cell ss:Index="6" ss:MergeAcross="1" ss:MergeDown="1" ss:StyleID="s8"><Data
      ss:Type="String">' || rec.number_agr || '</Data></Cell>
    <Cell ss:StyleID="s9"><Data ss:Type="String">Date </Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s8"><Data ss:Type="String">' || rec.doc_date || '</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">Пункт назначения</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s5"><Data ss:Type="String">Контракт</Data></Cell>
    <Cell ss:Index="8" ss:StyleID="s5"><Data ss:Type="String">Дата</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s9"><Data ss:Type="String">Steamer </Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s1"/>
    <Cell ss:Index="4" ss:StyleID="s9"><Data ss:Type="String">Letter of credit</Data></Cell>
    <Cell ss:Index="6" ss:MergeAcross="1" ss:MergeDown="1" ss:StyleID="s1"/>
    <Cell ss:StyleID="s9"><Data ss:Type="String">Date </Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s1"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">Теплоход (наименование)</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s5"><Data ss:Type="String">Аккредитив</Data></Cell>
    <Cell ss:Index="8" ss:StyleID="s5"><Data ss:Type="String">Дата</Data></Cell>
   </Row>';
end loop;
DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));  
dbms_output.put_line('9999999999999');
    
FOR rec IN Q_pril LOOP
    v_str:=v_str||rec.num_appl||'&#10;';
    v_str_date:=v_str_date||rec.date_APPL||'&#10;';
END LOOP;
dbms_output.put_line(v_str);  

  l_row_d := '  
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s9"><Data ss:Type="String">Bill of Lading</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s1"/>
    <Cell ss:Index="4" ss:StyleID="s9"><Data ss:Type="String">Appendix</Data></Cell>
    <Cell ss:Index="6" ss:MergeAcross="1" ss:MergeDown="9" ss:StyleID="s1"><Data ss:Type="String">' || v_str || '</Data></Cell>
    <Cell ss:StyleID="s9"><Data ss:Type="String">Date </Data></Cell>
    <Cell ss:MergeDown="9" ss:StyleID="s1"><Data ss:Type="String">' || v_str_date || '</Data></Cell>
   </Row>';
 DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));

  dbms_output.put_line('++++++++++ ');


for rec in Q_zag loop
 l_row_d := ' 
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">Коносамент</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s5"><Data ss:Type="String">Приложение</Data></Cell>
    <Cell ss:Index="8" ss:StyleID="s5"><Data ss:Type="String">Дата</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s9"><Data ss:Type="String">Freight car</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s1"><Data ss:Type="String"> </Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">Вагон</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s9"><Data ss:Type="String">Railway Bill</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s1"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">Ж.-д накладная </Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s9"><Data ss:Type="String">Hauling unit </Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s1"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">транспортное средство</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s9"><Data ss:Type="String">CMR</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s1"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">Накладная</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="11.25">
    <Cell ss:StyleID="s9"><Data ss:Type="String">Date of Shipment</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s8"><Data ss:Type="String">' || rec.doc_date || '</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s9"><Data ss:Type="String">Terms of Delivery</Data></Cell>
    <Cell ss:Index="6" ss:StyleID="s4"><Data ss:Type="String">' || rec.supply || '</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:StyleID="s5"><Data ss:Type="String">Дата отгрузки</Data></Cell>
    <Cell ss:Index="4" ss:StyleID="s5"><Data ss:Type="String">Условия поставки</Data></Cell>
   </Row> ';
 end loop;  
 
 DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5')); 
 --==================================================================================


 dbms_output.put_line('/\/\/\/\/\/\/\/\/\/\/\/\ ');

 l_row_d := ' 
 <Row ss:Index="37" ss:AutoFitHeight="0" ss:Height="30">
    <Cell ss:MergeAcross="1" ss:StyleID="s10"><Data ss:Type="String">Description of Goods</Data></Cell>
    <Cell ss:StyleID="s10"><Data ss:Type="String">Quantity of packing cases</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s10"><Data ss:Type="String">Weight  Net</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s10"><Data ss:Type="String">Weight   Gross</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s10"><Data ss:Type="String">Amount</Data></Cell>
   </Row>
  
   <Row ss:AutoFitHeight="0" ss:Height="18.75">
    <Cell ss:MergeAcross="1" ss:StyleID="s11"><Data ss:Type="String">Наименование товара</Data></Cell>
    <Cell ss:StyleID="s11"><Data ss:Type="String">Количество мест</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s11"><Data ss:Type="String">Вес   нетто</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s11"><Data ss:Type="String">Вес    брутто</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s11"><Data ss:Type="String">Сумма</Data></Cell>
   </Row>
    
   <Row ss:AutoFitHeight="0" ss:Height="18.75">
    <Cell ss:Index="3" ss:StyleID="s11"></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s11"><Data ss:Type="String">kg / кг</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s11"><Data ss:Type="String">kg / кг</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s11"><Data ss:Type="String">USD</Data></Cell>
   </Row>';
   DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));
   
   dbms_output.put_line(' ************ ');
   v_summa_ac:=0;
 for rec in Q0 loop

  l_row_d := '
   <Row ss:AutoFitHeight="0" ss:Height="29.25">
    <Cell ss:MergeAcross="1" ss:StyleID="s12"><Data ss:Type="String">' || rec.name_invoice || '</Data></Cell>';
   DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));

--    v_number_agr := rec.number_agr;
--    v_reg_date_agr := to_char(rec.reg_date_agr, 'dd.mm.yyyy');
    v_summa_ac := nvl(v_summa_ac,0) + nvl(rec.summa_ac,0);
    v_weight_nett := nvl(v_weight_nett, 0) + NVL(rec.weight_nett, 0);
    v_weight_gross := nvl(v_weight_gross, 0) + NVL(rec.weight_gross, 0);
    v_row_end := Q0%ROWCOUNT + 40;
 end loop;   

  l_row_d := '
    <Cell ss:StyleID="s12"><Data ss:Type="Number">' || v_sum_quan_fakt || '</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s12"><Data ss:Type="Number">' || round(v_weight_nett,3) || '</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s12"><Data ss:Type="Number">' || round(v_weight_gross,3) || '</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s12"><Data ss:Type="Number">' || round(v_summa_ac,2) || '</Data></Cell>
   </Row>';
   DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));
   
-- dbms_output.put_line('===========');  

 dbms_output.put_line('end loop invoice');
 
 
 l_row_d := '
    <Row ss:AutoFitHeight="0" ss:Height="10.5"></Row>  
	<Row ss:AutoFitHeight="0" ss:Height="10.5"></Row>  
   <Row ss:AutoFitHeight="0" ss:Height="10.5"></Row>  
   <Row ss:AutoFitHeight="0" ss:Height="10.5"></Row>
     
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:Index="5" ss:StyleID="s1"><Data ss:Type="String">TOTAL</Data></Cell>
    <Cell ss:Index="7" ss:MergeDown="1" ss:StyleID="s4"><Data ss:Type="String">USD</Data></Cell>
    <Cell ss:MergeAcross="1" ss:MergeDown="1" ss:StyleID="s7"><Data ss:Type="Number">' || round(v_summa_ac,2) || '</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:Index="5" ss:StyleID="s3"><Data ss:Type="String">ИТОГО</Data></Cell>
   </Row>
   
   <Row ss:AutoFitHeight="0" ss:Height="10.5"></Row>  
	<Row ss:AutoFitHeight="0" ss:Height="10.5"></Row>
	
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:Index="5" ss:StyleID="s1"><Data ss:Type="String">AMOUNT TO BE PAID</Data></Cell>
    <Cell ss:Index="7" ss:MergeDown="1" ss:StyleID="s4"><Data ss:Type="String">USD</Data></Cell>
    <Cell ss:MergeAcross="1" ss:MergeDown="1" ss:StyleID="s7"><Data ss:Type="Number">' || round(v_summa_ac,2) || '</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:Index="5" ss:StyleID="s3"><Data ss:Type="String">ВCЕГО ПО СЧЕТУ</Data></Cell>
	</Row>
	
	 <Row ss:AutoFitHeight="0" ss:Height="10.5"></Row>  
	<Row ss:AutoFitHeight="0" ss:Height="10.5"></Row>
    <Row ss:AutoFitHeight="0" ss:Height="10.5"></Row>  
	<Row ss:AutoFitHeight="0" ss:Height="10.5"></Row>
	
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:Index="5" ss:StyleID="s1"><Data ss:Type="String">' || naim_eng || '</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="10.5">
    <Cell ss:Index="5" ss:StyleID="s55"><Data ss:Type="String">' || naim_rus || '</Data></Cell>
   </Row>';  
 --==============================================================================
DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));

 dbms_output.put_line('+++++++++++');
 
 l_row_d := ' 
 </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.51181102362204722"/>
    <Footer x:Margin="0.51181102362204722"/>
    <PageMargins x:Bottom="0.27559055118110237" x:Left="0.19685039370078741"
     x:Right="0.19685039370078741" x:Top="0.35433070866141736"/>
   </PageSetup>
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <PageBreakZoom>60</PageBreakZoom>

   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>56</ActiveRow>
     <ActiveCol>1</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 
</Workbook> ';
DBMS_LOB.append(l_peremen, CONVERT(l_row_d, 'CL8ISO8859P5'));

  dbms_output.put_line('end sheet invoice ......... ');
  
 --КОНЕЦ ПРОЦЕДУРЫ 
-- формируем BLOB Excel файла для выгрузки

      l_file := utl_i18n.string_to_raw(substr(l_peremen, 1, 4000), 'AL32UTF8' /*'CL8ISO8859P5'*/ /*'CL8MSWIN1251'*/); 
      FOR i IN 2 .. ceil(dbms_lob.getlength(l_peremen) / 4000) LOOP
        dbms_lob.append(l_file,utl_i18n.string_to_raw(substr(l_peremen, (i - 1) * 4000 + 1, 4000), 'AL32UTF8' /*'CL8ISO8859P5'*/ /*'CL8MSWIN1251'*/));
      END LOOP;
 
-- записываем BLOB и CLOB Excel файла

    UPDATE product.pro_file_excel t SET t.blob_excel = l_file, t.file_excel = l_peremen
       WHERE t.name_excel = 'unload_xml_rlw'
         AND t.num_excel = 8;
      COMMIT;
      EXCEPTION WHEN OTHERS THEN
       NULL;
  
    dbms_output.put_line('end of end ');  
 
 
 
END;
