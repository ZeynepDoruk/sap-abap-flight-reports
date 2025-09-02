*--------------------------------------------------------------------*
* Report Name: Z_SFLIGHT_REPORT
* Purpose   : Display filtered flight reservation details
* Author    : Zeynep
* Date      : 10-12-2024
*--------------------------------------------------------------------*
REPORT z_sflight_report.
CLASS zcl_excel DEFINITION LOAD.
CLASS zcl_excel_worksheet DEFINITION LOAD.


*--------------------------------------------------------------------*
* Selection Screen for Filtering
SELECTION-SCREEN BEGIN OF BLOCK filters WITH FRAME TITLE TEXT-001.

PARAMETERS: p_carna  TYPE scarr-carrname,
            p_name   TYPE scustom-name,
            p_citfr  TYPE spfli-cityfrom,
            p_cityto TYPE spfli-cityto,
            p_fldafr TYPE sflight-fldate,
            p_fldato TYPE sflight-fldate MODIF ID d_r,
            p_class  TYPE sbook-class.  " Add Class Filter (Radio Button)

SELECTION-SCREEN END OF BLOCK filters.

DATA: it_result      TYPE TABLE OF zreservation_data, " Output Table
      g_layout       TYPE lvc_s_layo,
      gt_fieldcat    TYPE lvc_t_fcat,
      gt_sort        TYPE lvc_t_sort,
      wa_sort        TYPE lvc_s_sort,
      gt_top_of_page TYPE slis_t_listheader, " Top-of-page data
      gs_top_of_page TYPE slis_listheader.  " Single row structure for ALV header

FIELD-SYMBOLS: <wa_fields> TYPE lvc_s_fcat.

*--------------------------------------------------------------------*
* Start-of-selection event to fetch data and display report
*--------------------------------------------------------------------*
START-OF-SELECTION.         " Ekranı 0100'de çağırıyoruz
  IF sy-batch EQ 'X'. " Eğer job tarafından çalıştırılıyorsa
    PERFORM select_data.
    PERFORM sort.
    PERFORM gs_layout.
    PERFORM gs_field_catalog.
    PERFORM job_export_to_excel. " Excel'e otomatik export
    EXIT. " İşlemi bitir
  ELSE.
    PERFORM select_data.
    PERFORM sort.
    PERFORM gs_layout.
    PERFORM gs_field_catalog.
    PERFORM write. " Normal ALV ekranı gösterimi
  ENDIF.
*--------------------------------------------------------------------*
* Fetch data and join required tables
*--------------------------------------------------------------------*
FORM select_data.
  " Clear the result table before fetching new data
  CLEAR it_result.

  " Ensure parameters have values if not provided by user
  IF p_carna IS INITIAL.
    p_carna = '%'.
  ENDIF.

  IF p_citfr IS INITIAL.
    p_citfr = '%'.
  ENDIF.

  IF p_cityto IS INITIAL.
    p_cityto = '%'.
  ENDIF.

  IF p_fldafr IS INITIAL.
    p_fldafr = sy-datum . " Default start date if not provided
  ENDIF.

  IF p_fldato IS INITIAL.
    p_fldato = '9999-12-31'.  " Default end date if not provided
  ENDIF.

  DATA: lv_name_filter TYPE string.
  IF p_name IS NOT INITIAL.
    lv_name_filter = '%' && p_name && '%'.
  ELSE.
    lv_name_filter = '%'.  " No filtering by name
  ENDIF.
  IF p_class IS INITIAL.
    p_class = '%'.  " Varsayılan olarak tüm sınıflar
  ENDIF.

  DATA: lv_message TYPE string.

  lv_message = 'Carrier: '.
  IF p_carna <> '%'.
    lv_message = lv_message && p_carna.
  ELSE.
    lv_message = lv_message && 'All carriers'.
  ENDIF.

  lv_message = lv_message && ', CityFrom: '.
  IF p_citfr <> '%'.
    lv_message = lv_message && p_citfr.
  ELSE.
    lv_message = lv_message && 'All cities (Departure)'.
  ENDIF.

  lv_message = lv_message && ', CityTo: '.
  IF p_cityto <> '%'.
    lv_message = lv_message && p_cityto.
  ELSE.
    lv_message = lv_message && 'All cities (Destination)'.
  ENDIF.

  lv_message = lv_message && ', Flight Date: '.
  IF p_fldafr <> '9999-01-01' AND p_fldato <> '9999-12-31'.
    lv_message = lv_message && p_fldafr && ' to ' && p_fldato.
  ELSE.
    lv_message = lv_message && 'All dates'.
  ENDIF.

  lv_message = lv_message && ', Name: '.
  IF p_name IS NOT INITIAL.
    lv_message = lv_message && p_name.
  ELSE.
    lv_message = lv_message && 'No filter on name'.
  ENDIF.

  " Class bilgisi ekleniyor
  lv_message = lv_message && ', Class: '.
  IF p_class IS INITIAL OR p_class = '%'.
    lv_message = lv_message && 'All classes'.
  ELSE.
    lv_message = lv_message && p_class.
  ENDIF.
  "
  " MESSAGE lv_message TYPE 'I'.
  DATA result TYPE string.
  " Fetch data from the ZRESERVATION_DATA structure
  SELECT a~carrname,
         b~name AS name1,
         c~bookid,
         d~cityfrom,
         d~cityto,
         c~order_date,
         e~fldate,
         c~class,
         e~price,
         e~seatsmax,
         e~seatsocc,
         e~seatsmax_b,
         e~seatsocc_b,
         e~seatsmax_f,
         e~seatsocc_f,
         ( e~seatsmax + e~seatsmax_b + e~seatsmax_f ) AS seats_total,
         ( e~seatsocc + e~seatsocc_b + e~seatsocc_f ) AS seats_occ_total
    INTO CORRESPONDING FIELDS OF TABLE @it_result
    UP TO 200 ROWS
    FROM scarr AS a
    INNER JOIN spfli AS d ON a~carrid = d~carrid
    INNER JOIN sbook AS c ON d~connid = c~connid
    INNER JOIN sflight AS e ON d~connid = e~connid
    INNER JOIN scustom AS b ON c~customid = b~id.
*    WHERE a~carrname LIKE @p_carna
*       AND d~cityfrom LIKE @p_citfr
*       AND d~cityto LIKE @p_cityto
*       AND e~fldate BETWEEN @p_fldafr AND @p_fldato
*       AND e~fldate IS NOT NULL
*       AND e~seatsmax > 0
*       AND c~class LIKE @p_class
*       AND b~name LIKE @lv_name_filter.

  LOOP AT it_result INTO DATA(line).

    IF line-seats_occ_total = line-seats_total.
      line-occ_rate = '1'.
    ELSE.
      line-occ_rate = |{ line-seats_occ_total }/{ line-seats_total }|.
    ENDIF.

    " Güncellenen satırı tabloya geri yaz
    MODIFY it_result FROM line.

  " Tekrarlayan kayıtları sil
  SORT it_result BY carrname name1 bookid cityfrom cityto fldate class.
  DELETE ADJACENT DUPLICATES FROM it_result COMPARING carrname name1 bookid cityfrom cityto fldate class.

  " Son sıralama (örnek: Airline ve Doluluk Oranı)
  SORT it_result BY carrname DESCENDING occ_rate DESCENDING.

  ENDLOOP.
  DATA: lo_alv_grid    TYPE REF TO cl_gui_alv_grid,
        lo_container   TYPE REF TO cl_gui_custom_container,
        lo_grid_layout TYPE lvc_s_layo.

* ALV için container ve grid nesnelerini oluşturun
  CREATE OBJECT lo_container
    EXPORTING
      container_name = 'ALV_CONTAINER'. " Container adı (ekran üzerindeki yerleşim adı)

  CREATE OBJECT lo_alv_grid
    EXPORTING
      i_parent = lo_container.

* ALV için veri yapısını belirleyin
  lo_grid_layout-grid_title = 'Rezervation information '.
* ALV'nin ilk gösterimini yapın
  CALL METHOD lo_alv_grid->set_table_for_first_display
    EXPORTING
      i_structure_name = 'zreservation_data'
    CHANGING
      it_outtab        = it_result.




ENDFORM.

*--------------------------------------------------------------------*
* Sort the results based on multiple fields
*--------------------------------------------------------------------*
FORM sort.
  CLEAR: gt_sort, wa_sort.
  wa_sort-spos = '1'. " Sort by Airline Name
  wa_sort-fieldname = 'CARRNAME'.
  wa_sort-down = 'X'.
  APPEND wa_sort TO gt_sort.

  wa_sort-spos = '2'." Sort by Reservation Number
  wa_sort-fieldname = 'BOOKID'.
  wa_sort-down = 'X'.
  APPEND wa_sort TO gt_sort.

  wa_sort-spos = '3'. " Doluluk oranına göre sıralama
  wa_sort-fieldname = 'OCC_RATE'.
  wa_sort-down = 'X'.
  APPEND wa_sort TO gt_sort.

ENDFORM.

*--------------------------------------------------------------------*
* Generate Field Catalog for ALV
*--------------------------------------------------------------------*
FORM gs_field_catalog.
  CLEAR gt_fieldcat.

  " Generate field catalog for the structure ZRESERVATION_DATA
  CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
    EXPORTING
      i_structure_name       = 'ZRESERVATION_DATA'
    CHANGING
      ct_fieldcat            = gt_fieldcat
    EXCEPTIONS
      inconsistent_interface = 1
      program_error          = 2
      OTHERS                 = 3.

  IF sy-subrc <> 0.
    WRITE: / 'Field catalog could not be created! Error code:', sy-subrc.
    EXIT.
  ENDIF.

  " Assign display names for each column
  LOOP AT gt_fieldcat ASSIGNING <wa_fields>.
    CASE <wa_fields>-fieldname.
      WHEN 'CARRNAME'.
        <wa_fields>-coltext = 'Airline Name'.
      WHEN 'NAME1'.
        <wa_fields>-coltext = 'Customer Name'.
      WHEN 'BOOKID'.
        <wa_fields>-coltext = 'Reservation Number'.
      WHEN 'CITYFROM'.
        <wa_fields>-coltext = 'Departure City'.
      WHEN 'CITYTO'.
        <wa_fields>-coltext = 'Arrival City'.
      WHEN 'ORDER_DATE'.
        <wa_fields>-coltext = 'Booking Date'.
      WHEN 'FLDATE'.
        <wa_fields>-coltext = 'Flight Date'.
      WHEN 'PRICE'.
        <wa_fields>-coltext = 'Price'.
      WHEN 'SEATSMAX'.
        <wa_fields>-coltext = 'Max Seats Economy'.
      WHEN 'SEATS_OCC'.
        <wa_fields>-coltext = 'Seats Occupied Economy'.
      WHEN 'SEATSMAX_B'.
        <wa_fields>-coltext = 'Max Seats Business'.
      WHEN 'SEATS_OCC_B'.
        <wa_fields>-coltext = 'Seats Occupied Business'.
      WHEN 'SEATSMAX_F'.
        <wa_fields>-coltext = 'Max Seats First'.
      WHEN 'SEATS_OCC_F'.
        <wa_fields>-coltext = 'Seats Occupied First'.
      WHEN 'SEATS_TOTAL'.
        <wa_fields>-coltext = 'Toplam Koltuk Sayısı'.
      WHEN 'SEATS_OCC_TOTAL'.
        <wa_fields>-coltext = 'Toplam Dolu Koltuk Sayısı'.
      WHEN 'OCC_RATE'.
        <wa_fields>-coltext = 'Doluluk Oranı (%)'.
      WHEN OTHERS.
        <wa_fields>-coltext = <wa_fields>-fieldname.  " Default to field name if not mapped
    ENDCASE.
  ENDLOOP.
ENDFORM.

*--------------------------------------------------------------------*
* Define ALV layout options
*--------------------------------------------------------------------*
FORM gs_layout.
  CLEAR g_layout.
  g_layout-cwidth_opt = 'X'.  " Column width optimization
  g_layout-zebra = 'X'.       " Zebra striping
  g_layout-edit = 'X'.        " Enable editing of rows
ENDFORM.

*--------------------------------------------------------------------*
* Write ALV output
*--------------------------------------------------------------------*
FORM write.

  CLEAR gt_top_of_page.
  gs_top_of_page = 'Flight Report'.
  APPEND gs_top_of_page TO gt_top_of_page.

  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY_LVC'
    EXPORTING
      i_callback_program       = sy-repid
      is_layout_lvc            = g_layout
      it_sort_lvc              = gt_sort
      it_fieldcat_lvc          = gt_fieldcat
      i_callback_pf_status_set = 'SET_PF_STATUS_EXCEL'
      i_callback_user_command  = 'USER_COMMAND'
      i_callback_top_of_page   = 'TOP_OF_PAGE'
    TABLES
      t_outtab                 = it_result
    EXCEPTIONS
      program_error            = 1
      OTHERS                   = 2.

  IF sy-subrc <> 0.
    WRITE: / 'ALV Grid display failed. Error code:', sy-subrc.
  ENDIF.
ENDFORM.
*---------------------------------------------------------------------*
* COMMENT_BUILD: Header ve Alt Bilgiler Hazırlanır
*---------------------------------------------------------------------*
FORM comment_build USING it_top_of_page TYPE slis_t_listheader.
  DATA: ls_line TYPE slis_listheader.

  " Ana başlık
  CLEAR ls_line.
  ls_line-typ  = 'H'.
  ls_line-info = 'Flight Reservation Report'.
  APPEND ls_line TO it_top_of_page.

  " Alt başlık
  CLEAR ls_line.
  ls_line-typ = 'S'.
  ls_line-key = 'Tarih: '.
  WRITE sy-datum TO ls_line-info.
  APPEND ls_line TO it_top_of_page.

  CLEAR ls_line.
  ls_line-typ = 'S'.
  ls_line-key = 'Kullanıcı: '.
  WRITE sy-uname TO ls_line-info.
  APPEND ls_line TO it_top_of_page.

  CLEAR ls_line.
  ls_line-typ = 'S'.
  ls_line-key = 'Not: '.
  ls_line-info = 'Doluluk oranı 1 olması uçağın full dolu olduğu anlamına gelir'.
  APPEND ls_line TO it_top_of_page.

ENDFORM.
*--------------------------------------------------------------------*
* Export to Excel
*--------------------------------------------------------------------*
*--------------------------------------------------------------------*
* Export to Excel
*--------------------------------------------------------------------*
*--------------------------------------------------------------------*
* Form Export to Excel
*--------------------------------------------------------------------*
FORM export_to_excel.
  TRY.
      " File dialog initialization
      DATA: lv_path     TYPE string,
            lv_filename TYPE string,
            lv_fullpath TYPE string.

      lv_filename = 'Reservation_Report.xlsx'. " Default filename for Excel

      CALL METHOD cl_gui_frontend_services=>file_save_dialog
        EXPORTING
          default_extension = 'xlsx'
          default_file_name = lv_filename
        CHANGING
          filename          = lv_filename
          path              = lv_path
          fullpath          = lv_fullpath
        EXCEPTIONS
          OTHERS            = 1.

      IF sy-subrc = 0 AND lv_fullpath IS NOT INITIAL.

        " Create Excel file using ABAP2XLSX
        DATA o_xl TYPE REF TO zcl_excel.
        DATA(o_converter) = NEW zcl_excel_converter( ).

        o_converter->convert(
          EXPORTING
            it_table = it_result " Input internal table
          CHANGING
            co_excel = o_xl ).

        " Get active worksheet
        DATA(o_xl_ws) = o_xl->get_active_worksheet( ).
        o_xl_ws->freeze_panes( ip_num_rows = 1 ). " Freeze first row

        " Generate xstring
        DATA(o_xl_writer) = CAST zif_excel_writer( NEW zcl_excel_writer_2007( ) ).
        DATA(lv_xl_xdata) = o_xl_writer->write_file( o_xl ).

        " Convert xstring to binary table
        DATA(it_raw_data) = cl_bcs_convert=>xstring_to_solix(
          EXPORTING
            iv_xstring = lv_xl_xdata ).

        " Save file to frontend
        cl_gui_frontend_services=>gui_download(
          EXPORTING
            bin_filesize = 0
            filename     = lv_fullpath
            filetype     = 'BIN'
          CHANGING
            data_tab     = it_raw_data ).

        MESSAGE 'Excel dosyası başarıyla kaydedildi!' TYPE 'I'.

      ELSE.
        MESSAGE 'Dosya kaydedilmedi!' TYPE 'W'.
      ENDIF.

    CATCH cx_root INTO DATA(lx_error).
      MESSAGE lx_error->get_text( ) TYPE 'E'.
  ENDTRY.

ENDFORM.

*--------------------------------------------------------------------*
* Form USER_COMMAND
*--------------------------------------------------------------------*
FORM user_command USING ucomm LIKE sy-ucomm selfield TYPE slis_selfield.
  CASE ucomm.
    WHEN 'EXCEL'.
      PERFORM export_to_excel.
    WHEN 'BACK'.
      CLEAR : it_result, p_carna, p_name, p_citfr, p_cityto, p_fldafr, p_fldato, p_class.
      CLEAR : gt_top_of_page.
      LEAVE TO SCREEN 0. " Başlangıç ekranına dönün
    WHEN OTHERS.
      MESSAGE 'Bilinmeyen bir komut girildi!' TYPE 'I'.
  ENDCASE.
ENDFORM.

*--------------------------------------------------------------------*
* Form SET_PF_STATUS_EXCEL
*--------------------------------------------------------------------*
FORM set_pf_status_excel USING rt_extab TYPE slis_t_extab.
  SET PF-STATUS '0100'. " Custom GUI status with Excel button
  " Add Excel button to the toolbar extension table
  APPEND 'EXCEL' TO rt_extab.
  APPEND 'BACK' TO rt_extab.
ENDFORM.



*--------------------------------------------------------------------*
* Top of Page Event for ALV Report
*--------------------------------------------------------------------*
FORM top_of_page.
  DATA: l_document         TYPE REF TO cl_dd_document,
        it_list_commentary TYPE slis_t_listheader,
        ls_line            TYPE slis_listheader.

  PERFORM comment_build USING it_list_commentary.

  CREATE OBJECT l_document.
  CALL METHOD l_document->add_text
    EXPORTING
      text = 'Flight Reservation Report - Generated by ALV'.

  EXPORT it_list_commentary FROM it_list_commentary
    TO MEMORY ID 'DYNDOS_FOR_ALV'.

  CALL FUNCTION 'REUSE_ALV_GRID_COMMENTARY_SET'
    EXPORTING
      document = l_document
      bottom   = space.
  CALL FUNCTION 'REUSE_ALV_COMMENTARY_WRITE'
    EXPORTING
      i_logo             = 'SAP_LOGO'
      it_list_commentary = it_list_commentary.
ENDFORM.



FORM job_export_to_excel.
  TRY.

        " Create Excel file using ABAP2XLSX
        DATA o_xl TYPE REF TO zcl_excel.
        DATA(o_converter) = NEW zcl_excel_converter( ).

        o_converter->convert(
          EXPORTING
            it_table = it_result " Input internal table
          CHANGING
            co_excel = o_xl ).

        " Get active worksheet
        DATA(o_xl_ws) = o_xl->get_active_worksheet( ).
        o_xl_ws->freeze_panes( ip_num_rows = 1 ). " Freeze first row

        " Generate xstring
        DATA(o_xl_writer) = CAST zif_excel_writer( NEW zcl_excel_writer_2007( ) ).
        DATA(lv_xl_xdata) = o_xl_writer->write_file( o_xl ).

        " Convert xstring to binary table
        DATA(it_raw_data) = cl_bcs_convert=>xstring_to_solix(
          EXPORTING
            iv_xstring = lv_xl_xdata ).

        " Save file to frontend
        cl_gui_frontend_services=>gui_download(
          EXPORTING
            bin_filesize = 0
            filename     = 'C:\Users\Zeynep\Desktop\flight_report.xlsx'
            filetype     = 'BIN'
          CHANGING
            data_tab     = it_raw_data ).

        MESSAGE 'Excel dosyası başarıyla kaydedildi!' TYPE 'I'.



    CATCH cx_root INTO DATA(lx_error).
      MESSAGE lx_error->get_text( ) TYPE 'E'.
  ENDTRY.
ENDFORM.
