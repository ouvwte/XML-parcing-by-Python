CREATE OR REPLACE PROCEDURE PRO_UNLOAD_XML_RLW IS
    l_peremen    CLOB;      -- итоговый XML (CLOB)
    l_piece      CLOB;      -- временный кусок
    l_file_blob  BLOB;      -- итоговый BLOB
    v_warning    INTEGER := 0;

    v_year       NUMBER := 2025; -- при необходимости сделать параметром
    v_num_excel  NUMBER := 8;
    v_name_excel VARCHAR2(200) := 'unload_xml_rlw';

    -- массив названий месяцев
    TYPE t_months IS VARRAY(12) OF VARCHAR2(20);
    months t_months := t_months(
      'январь','февраль','март','апрель','май','июнь',
      'июль','август','сентябрь','октябрь','ноябрь','декабрь'
    );

    ----------------------------------------------------------------
    -- число дней в месяце
    ----------------------------------------------------------------
    FUNCTION days_count(p_year NUMBER, p_month NUMBER) RETURN NUMBER IS
      v_date DATE;
    BEGIN
      v_date := TO_DATE(p_year || '-' || p_month || '-01','YYYY-MM-DD');
      RETURN TO_NUMBER(TO_CHAR(LAST_DAY(v_date),'DD'));
    END;

    ----------------------------------------------------------------
    -- получить значение для ячейки (plan/fact/dev/misc)
    ----------------------------------------------------------------
    FUNCTION get_val(p_year NUMBER, p_month NUMBER, p_day NUMBER, p_org_id NUMBER, p_field VARCHAR2) RETURN VARCHAR2 IS
      v_res VARCHAR2(4000);
    BEGIN
      CASE LOWER(p_field)
      WHEN 'plan' THEN
        SELECT NVL(TO_CHAR(plan_val), '') INTO v_res
          FROM railway.rlw_board
         WHERE god = p_year AND mes = p_month AND den = p_day AND org_id = p_org_id
         AND ROWNUM = 1;
      WHEN 'fact' THEN
        SELECT NVL(TO_CHAR(fact_val), '') INTO v_res
          FROM railway.rlw_board
         WHERE god = p_year AND mes = p_month AND den = p_day AND org_id = p_org_id
         AND ROWNUM = 1;
      WHEN 'dev' THEN
        SELECT NVL(TO_CHAR(dev_val), '') INTO v_res
          FROM railway.rlw_board
         WHERE god = p_year AND mes = p_month AND den = p_day AND org_id = p_org_id
         AND ROWNUM = 1;
      WHEN 'misc' THEN
        SELECT NVL(TO_CHAR(misc_val), '') INTO v_res
          FROM railway.rlw_board
         WHERE god = p_year AND mes = p_month AND den = p_day AND org_id = p_org_id
         AND ROWNUM = 1;
      ELSE
        v_res := '';
      END CASE;
      RETURN v_res;
    EXCEPTION
      WHEN NO_DATA_FOUND THEN RETURN '';
      WHEN OTHERS THEN RETURN '';
    END;

BEGIN
  ----------------------------------------------------------------
  -- очистка blob в product.pro_file_excel
  ----------------------------------------------------------------
  BEGIN
    UPDATE product.pro_file_excel
       SET blob_excel = EMPTY_BLOB()
     WHERE name_excel = v_name_excel
       AND num_excel = v_num_excel;
    COMMIT;
  EXCEPTION WHEN OTHERS THEN NULL;
  END;

  EXECUTE IMMEDIATE 'ALTER SESSION SET NLS_NUMERIC_CHARACTERS = ''. ''';

  ----------------------------------------------------------------
  -- Шапка XML
  ----------------------------------------------------------------
  l_piece := '<?xml version="1.0" encoding="UTF-8"?>' || CHR(10)
          || '<?mso-application progid="Excel.Sheet"?>' || CHR(10)
          || '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"'
          || ' xmlns:o="urn:schemas-microsoft-com:office:office"'
          || ' xmlns:x="urn:schemas-microsoft-com:office:excel"'
          || ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"'
          || ' xmlns:html="http://www.w3.org/TR/REC-html40">' || CHR(10)
          || ' <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">'
          || '  <Created>' || TO_CHAR(SYSDATE,'YYYY-MM-DD"T"HH24:MI:SS"Z"') || '</Created>'
          || '  <LastSaved>' || TO_CHAR(SYSDATE,'YYYY-MM-DD"T"HH24:MI:SS"Z"') || '</LastSaved>'
          || '  <Version>16.00</Version>'
          || ' </DocumentProperties>'
          || ' <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">'
          || '  <AllowPNG/>'
          || '  <RemovePersonalInformation/>'
          || ' </OfficeDocumentSettings>'
          || ' <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">'
          || '  <WindowHeight>12300</WindowHeight>'
          || '  <WindowWidth>28800</WindowWidth>'
          || '  <WindowTopX>0</WindowTopX>'
          || '  <WindowTopY>0</WindowTopY>'
          || '  <RefModeR1C1/>'
          || '  <ProtectStructure>False</ProtectStructure>'
          || '  <ProtectWindows>False</ProtectWindows>'
          || ' </ExcelWorkbook>';
  l_peremen := l_piece;

  ----------------------------------------------------------------
  -- Стили (полный блок из sql_exmp.sql)
  ----------------------------------------------------------------
  l_piece := '
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
  DBMS_LOB.APPEND(l_peremen, l_piece);

  ----------------------------------------------------------------
  -- Основной цикл по месяцам (январь..декабрь)
  ----------------------------------------------------------------
  FOR m IN 1..12 LOOP
    DECLARE
      v_days NUMBER := days_count(v_year, m);
      v_month_name VARCHAR2(30) := months(m);
      v_row_idx NUMBER := 1;
    BEGIN
      -- открываем Worksheet
      l_piece := '<Worksheet ss:Name="' || v_month_name || '">' || CHR(10)
              || ' <Table x:FullColumns="1" x:FullRows="1" ss:DefaultRowHeight="15">' || CHR(10);
      DBMS_LOB.APPEND(l_peremen, l_piece);

      -- колонки
      DBMS_LOB.APPEND(l_peremen, '   <Column ss:Width="30"/>' || CHR(10));
      DBMS_LOB.APPEND(l_peremen, '   <Column ss:Width="200"/>' || CHR(10));
      DBMS_LOB.APPEND(l_peremen, '   <Column ss:Width="30"/>' || CHR(10));
      FOR d IN 1..v_days LOOP
        DBMS_LOB.APPEND(l_peremen, '   <Column ss:Width="60"/>' || CHR(10));
        DBMS_LOB.APPEND(l_peremen, '   <Column ss:Width="60"/>' || CHR(10));
        DBMS_LOB.APPEND(l_peremen, '   <Column ss:Width="60"/>' || CHR(10));
        DBMS_LOB.APPEND(l_peremen, '   <Column ss:Width="40"/>' || CHR(10));
      END LOOP;

      -- строка заголовков
      DBMS_LOB.APPEND(l_peremen, '   <Row>' || CHR(10)
        || '    <Cell><Data ss:Type="String">#</Data></Cell>' || CHR(10)
        || '    <Cell><Data ss:Type="String">Юрлицо</Data></Cell>' || CHR(10)
        || '    <Cell><Data ss:Type="String">Прим.</Data></Cell>' || CHR(10));
      FOR d IN 1..v_days LOOP
        DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="String">d' || LPAD(d,2,'0') || ' plan</Data></Cell>' || CHR(10));
        DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="String">d' || LPAD(d,2,'0') || ' fact</Data></Cell>' || CHR(10));
        DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="String">d' || LPAD(d,2,'0') || ' dev</Data></Cell>' || CHR(10));
        DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="String">d' || LPAD(d,2,'0') || ' *</Data></Cell>' || CHR(10));
      END LOOP;
      DBMS_LOB.APPEND(l_peremen, '   </Row>' || CHR(10));

      -- строки по юрлицам
      FOR rec IN (SELECT DISTINCT org_id, org_name
                    FROM railway.rlw_board
                   WHERE god = v_year AND mes = m
                   ORDER BY org_name) LOOP
        v_row_idx := v_row_idx + 1;
        DBMS_LOB.APPEND(l_peremen, '   <Row>' || CHR(10));
        DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="Number">' || (v_row_idx-1) || '</Data></Cell>' || CHR(10));
        DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="String">' 
            || REPLACE(REPLACE(NVL(rec.org_name,''), '&','&amp;'), '<','&lt;') 
            || '</Data></Cell>' || CHR(10));
        DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="String"/></Cell>' || CHR(10));
        FOR d IN 1..v_days LOOP
          DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="String">' || get_val(v_year, m, d, rec.org_id, 'plan') || '</Data></Cell>' || CHR(10));
          DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="String">' || get_val(v_year, m, d, rec.org_id, 'fact') || '</Data></Cell>' || CHR(10));
          DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="String">' || get_val(v_year, m, d, rec.org_id, 'dev')  || '</Data></Cell>' || CHR(10));
          DBMS_LOB.APPEND(l_peremen, '    <Cell><Data ss:Type="String">' || get_val(v_year, m, d, rec.org_id, 'misc') || '</Data></Cell>' || CHR(10));
        END LOOP;
        DBMS_LOB.APPEND(l_peremen, '   </Row>' || CHR(10));
      END LOOP;

      -- сюда можно вставить справочную информацию

      -- закрываем Worksheet
      DBMS_LOB.APPEND(l_peremen, ' </Table>' || CHR(10) || '</Worksheet>' || CHR(10));
    END;
  END LOOP;

  ----------------------------------------------------------------
  -- Закрытие Workbook
  ----------------------------------------------------------------
  DBMS_LOB.APPEND(l_peremen, '</Workbook>');

  ----------------------------------------------------------------
  -- CLOB -> BLOB
  ----------------------------------------------------------------
  DBMS_LOB.CREATETEMPORARY(l_file_blob, TRUE, DBMS_LOB.SESSION);
  DBMS_LOB.CONVERTTOBLOB(
      dest_lob    => l_file_blob,
      src_clob    => l_peremen,
      amount      => DBMS_LOB.GETLENGTH(l_peremen),
      dest_offset => 1,
      src_offset  => 1,
      blob_csid   => 873,
      lang_context=> 0,
      warning     => v_warning
  );

  UPDATE product.pro_file_excel
     SET blob_excel = l_file_blob
   WHERE name_excel = v_name_excel
     AND num_excel = v_num_excel;
  COMMIT;
END PRO_UNLOAD_XML_RLW;
