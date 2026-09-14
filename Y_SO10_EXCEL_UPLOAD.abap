*&---------------------------------------------------------------------*
*& Report Y_SO10_EXCEL_UPLOAD
*&---------------------------------------------------------------------*
*& UPLOAD so10 texts from Excel file
*&---------------------------------------------------------------------*
REPORT y_so10_excel_update.

PARAMETERS:
  p_file  TYPE string LOWER CASE OBLIGATORY,
  p_langu TYPE sylangu DEFAULT sy-langu OBLIGATORY,
  p_test  AS CHECKBOX DEFAULT 'X'.

SELECTION-SCREEN PUSHBUTTON /1(20) btn_tpl USER-COMMAND tpl.

CLASS lcl_application DEFINITION FINAL.
  PUBLIC SECTION.
    METHODS:
      run,
      f4_file,
      download_template.

  PRIVATE SECTION.
    CONSTANTS:
      mc_tdobject TYPE tdobject VALUE 'TEXT',
      mc_tdid     TYPE tdid     VALUE 'ST'.

    DATA:
      BEGIN OF mc_status,
        updated TYPE char12   VALUE 'UPDATED',
        error   TYPE char12   VALUE 'ERROR',
        check   TYPE char12   VALUE 'CHECK_OK',
      END OF  mc_status.

    TYPES:
      BEGIN OF ty_excel,
        name    TYPE string,
        content TYPE string,
      END OF ty_excel,
      tt_excel TYPE STANDARD TABLE OF ty_excel WITH EMPTY KEY,
      BEGIN OF ty_log,
        row     TYPE i,
        name    TYPE tdobname,
        status  TYPE char12,
        message TYPE string,
        color   TYPE lvc_t_scol,
      END OF ty_log,
      tt_log TYPE STANDARD TABLE OF ty_log WITH EMPTY KEY.

    DATA:
      mt_excel   TYPE tt_excel,
      mt_log     TYPE tt_log,
      mv_total   TYPE i,
      mv_updated TYPE i,
      mv_checked TYPE i,
      mv_errors  TYPE i.

    METHODS:
      upload_excel,
      process_excel,
      process_text
        IMPORTING
          iv_row     TYPE i
          iv_name    TYPE string
          iv_content TYPE string,
      check_text_exists
        IMPORTING iv_name          TYPE tdobname
        RETURNING VALUE(rv_exists) TYPE abap_bool,
      read_text
        IMPORTING iv_name   TYPE tdobname
        EXPORTING
                  es_header TYPE thead
                  et_lines  TYPE STANDARD TABLE,
      save_text
        IMPORTING
                  is_header         TYPE thead
                  it_lines          TYPE STANDARD TABLE
        RETURNING VALUE(rv_success) TYPE abap_bool,
      content_to_tline
        IMPORTING iv_content      TYPE string
        RETURNING VALUE(rt_lines) TYPE tlinet,
      compare_text
        IMPORTING
                  it_expected     TYPE tlinet
                  it_actual       TYPE tlinet
        RETURNING VALUE(rv_equal) TYPE abap_bool,
      add_log
        IMPORTING
          iv_row     TYPE i
          iv_name    TYPE tdobname
          iv_status  TYPE char12
          iv_message TYPE string,
      update_statistics
        IMPORTING iv_status TYPE char12,
      display_log,
      set_column_texts
        IMPORTING io_columns TYPE REF TO cl_salv_columns_table,
      set_top_of_list
        IMPORTING io_alv TYPE REF TO cl_salv_table.
ENDCLASS.

CLASS lcl_application IMPLEMENTATION.

  METHOD run.
    upload_excel( ).
    process_excel( ).
    display_log( ).
  ENDMETHOD.

  METHOD f4_file.
    DATA:
      lt_file_table TYPE filetable,
      lv_rc         TYPE i,
      lv_action     TYPE i.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING
        window_title      = 'Выберите Excel файл'
        default_extension = 'xlsx'
        file_filter       = 'Excel (*.xlsx)|*.xlsx|'
      CHANGING
        file_table        = lt_file_table
        rc                = lv_rc
        user_action       = lv_action
      EXCEPTIONS
        OTHERS = 1 ).

    IF sy-subrc <> 0 OR lv_action <> cl_gui_frontend_services=>action_ok.
      RETURN.
    ENDIF.

    READ TABLE lt_file_table INDEX 1 INTO DATA(ls_file).
    IF sy-subrc = 0.
      p_file = ls_file-filename.
    ENDIF.
  ENDMETHOD.

  METHOD upload_excel.
    DATA:
      lv_xstring    TYPE xstring,
      lt_raw        TYPE solix_tab,
      lv_size       TYPE i,
      lo_excel      TYPE REF TO cl_fdt_xl_spreadsheet,
      lt_worksheets TYPE if_fdt_doc_spreadsheet=>t_worksheet_names,
      lv_worksheet  TYPE string,
      lr_data       TYPE REF TO data.

    cl_gui_frontend_services=>gui_upload(
      EXPORTING
        filename   = p_file
        filetype   = 'BIN'
      IMPORTING
        filelength = lv_size
      CHANGING
        data_tab   = lt_raw
      EXCEPTIONS
        OTHERS = 1 ).

    IF sy-subrc <> 0.
      MESSAGE 'Ошибка загрузки Excel файла' TYPE 'E'.
    ENDIF.

    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = lv_size
      IMPORTING
        buffer       = lv_xstring
      TABLES
        binary_tab   = lt_raw
      EXCEPTIONS
        OTHERS       = 1.

    IF sy-subrc <> 0.
      MESSAGE 'Ошибка преобразования Excel файла' TYPE 'E'.
    ENDIF.

    TRY.
        lo_excel = NEW cl_fdt_xl_spreadsheet(
          document_name = p_file
          xdocument     = lv_xstring ).
      CATCH cx_fdt_excel_core INTO DATA(lx_excel).
        MESSAGE lx_excel->get_text( ) TYPE 'E'.
    ENDTRY.

    lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names( IMPORTING worksheet_names = lt_worksheets ).

    IF lt_worksheets IS INITIAL.
      MESSAGE 'Excel файл не содержит листов' TYPE 'E'.
    ENDIF.

    READ TABLE lt_worksheets INDEX 1 INTO lv_worksheet.
    IF sy-subrc <> 0.
      MESSAGE 'Не удалось определить лист Excel' TYPE 'E'.
    ENDIF.

    TRY.
        lr_data = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
          lv_worksheet ).
      CATCH cx_fdt_excel_core INTO lx_excel.
        MESSAGE lx_excel->get_text( ) TYPE 'E'.
    ENDTRY.

    FIELD-SYMBOLS:
      <lt_excel>   TYPE STANDARD TABLE,
      <ls_excel>   TYPE any,
      <lv_name>    TYPE any,
      <lv_content> TYPE any.

    ASSIGN lr_data->* TO <lt_excel>.
    IF <lt_excel> IS NOT ASSIGNED.
      MESSAGE 'Не удалось получить данные Excel' TYPE 'E'.
    ENDIF.

    LOOP AT <lt_excel> ASSIGNING <ls_excel>.
      DATA(lv_row) = sy-tabix.

      IF lv_row = 1.
        CONTINUE.
      ENDIF.

      UNASSIGN:
        <lv_name>,
        <lv_content>.

      ASSIGN COMPONENT 1 OF STRUCTURE <ls_excel> TO <lv_name>.
      ASSIGN COMPONENT 2 OF STRUCTURE <ls_excel> TO <lv_content>.

      IF <lv_name> IS NOT ASSIGNED OR <lv_content> IS NOT ASSIGNED.
        CONTINUE.
      ENDIF.

      APPEND VALUE #(
        name    = CONV string( <lv_name> )
        content = CONV string( <lv_content> ) ) TO mt_excel.
    ENDLOOP.

    IF mt_excel IS INITIAL.
      MESSAGE 'Excel не содержит данных для обработки' TYPE 'E'.
    ENDIF.

    mv_total = lines( mt_excel ).
  ENDMETHOD.

  METHOD process_excel.
    LOOP AT mt_excel INTO DATA(ls_excel).
      process_text(
        iv_row     = sy-tabix + 1
        iv_name    = ls_excel-name
        iv_content = ls_excel-content ).
    ENDLOOP.
  ENDMETHOD.

  METHOD process_text.
    DATA:
      lv_name        TYPE tdobname,
      ls_header      TYPE thead,
      ls_read_header TYPE thead,
      lt_new_lines   TYPE STANDARD TABLE OF tline,
      lt_saved_lines TYPE STANDARD TABLE OF tline.

    lv_name = iv_name.
    CONDENSE lv_name.

    IF lv_name IS INITIAL.
      add_log(
        iv_row     = iv_row
        iv_name    = ''
        iv_status  = mc_status-error
        iv_message = 'Пустое имя SO10 текста' ).
      RETURN.
    ENDIF.

    IF iv_content IS INITIAL.
      add_log(
        iv_row     = iv_row
        iv_name    = lv_name
        iv_status  = mc_status-error
        iv_message = 'Пустое значение столбца "Содержимое"' ).
      RETURN.
    ENDIF.

    IF check_text_exists( lv_name ) = abap_false.
      add_log(
        iv_row     = iv_row
        iv_name    = lv_name
        iv_status  = mc_status-error
        iv_message = 'SO10 текст не найден' ).
      RETURN.
    ENDIF.

    IF p_test = abap_true.
      add_log(
        iv_row     = iv_row
        iv_name    = lv_name
        iv_status  = mc_status-check
        iv_message = 'Текст найден. Изменение не выполнялось' ).
      RETURN.
    ENDIF.

    lt_new_lines = content_to_tline( iv_content ).

    IF lt_new_lines IS INITIAL.
      add_log(
        iv_row     = iv_row
        iv_name    = lv_name
        iv_status  = mc_status-error
        iv_message = 'Не удалось преобразовать содержимое в TLINE' ).
      RETURN.
    ENDIF.

    ls_header = VALUE #(
      tdobject = mc_tdobject
      tdname   = lv_name
      tdid     = mc_tdid
      tdspras  = p_langu ).

    IF save_text(
         is_header = ls_header
         it_lines  = lt_new_lines ) = abap_false.

      add_log(
        iv_row     = iv_row
        iv_name    = lv_name
        iv_status  = mc_status-error
        iv_message = |SAVE_TEXT завершился ошибкой| ).

      ROLLBACK WORK.
      RETURN.
    ENDIF.

    read_text(
      EXPORTING
        iv_name   = lv_name
      IMPORTING
        es_header = ls_read_header
        et_lines  = lt_saved_lines ).

    IF lt_saved_lines IS INITIAL.
      add_log(
        iv_row     = iv_row
        iv_name    = lv_name
        iv_status  = mc_status-error
        iv_message = 'SAVE_TEXT выполнен, но READ_TEXT не вернул содержимое' ).

      ROLLBACK WORK.
      RETURN.
    ENDIF.

    IF compare_text(
         it_expected = lt_new_lines
         it_actual   = lt_saved_lines ) = abap_false.

      add_log(
        iv_row     = iv_row
        iv_name    = lv_name
        iv_status  = mc_status-error
        iv_message = 'Пост-проверка READ_TEXT: содержимое не совпадает' ).

      ROLLBACK WORK.
      RETURN.
    ENDIF.

    COMMIT WORK AND WAIT.

    add_log(
      iv_row     = iv_row
      iv_name    = lv_name
      iv_status  = mc_status-updated
      iv_message = 'Текст успешно перезаписан и проверен через READ_TEXT' ).
  ENDMETHOD.

  METHOD check_text_exists.
    DATA:
      ls_header TYPE thead,
      lt_lines  TYPE STANDARD TABLE OF tline.

    rv_exists = abap_false.

    CALL FUNCTION 'READ_TEXT'
      EXPORTING
        id        = mc_tdid
        language  = p_langu
        name      = iv_name
        object    = mc_tdobject
      IMPORTING
        header    = ls_header
      TABLES
        lines     = lt_lines
      EXCEPTIONS
        id        = 1
        language  = 2
        name      = 3
        not_found = 4
        object    = 5
        OTHERS    = 6.

    rv_exists = xsdbool( sy-subrc = 0 ).
  ENDMETHOD.

  METHOD read_text.
    CLEAR:
      es_header,
      et_lines.

    CALL FUNCTION 'READ_TEXT'
      EXPORTING
        id        = mc_tdid
        language  = p_langu
        name      = iv_name
        object    = mc_tdobject
      IMPORTING
        header    = es_header
      TABLES
        lines     = et_lines
      EXCEPTIONS
        id        = 1
        language  = 2
        name      = 3
        not_found = 4
        object    = 5
        OTHERS    = 6.
  ENDMETHOD.

  METHOD save_text.
    CALL FUNCTION 'SAVE_TEXT'
      EXPORTING
        header          = is_header
        savemode_direct = 'X'
      TABLES
        lines           = it_lines
      EXCEPTIONS
        id              = 1
        language        = 2
        name            = 3
        object          = 4
        OTHERS          = 5.

    rv_success = xsdbool( sy-subrc = 0 ).
  ENDMETHOD.

  METHOD content_to_tline.
    DATA:
      lt_text    TYPE STANDARD TABLE OF string,
      lv_content TYPE string.

    lv_content = iv_content.

    REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf
      IN lv_content WITH cl_abap_char_utilities=>newline.

    SPLIT lv_content AT cl_abap_char_utilities=>newline INTO TABLE lt_text.

    LOOP AT lt_text INTO DATA(lv_line).
      DATA(lv_length) = strlen( lv_line ).

      IF lv_length = 0.
        APPEND VALUE #( tdformat = '*' tdline = '' ) TO rt_lines.
        CONTINUE.
      ENDIF.

      DATA(lv_offset) = 0.

      WHILE lv_offset < lv_length.
        DATA(lv_remaining) = lv_length - lv_offset.
        DATA(lv_chunk) = COND string(
          WHEN lv_remaining > 132
          THEN lv_line+lv_offset(132)
          ELSE lv_line+lv_offset(lv_remaining) ).

        APPEND VALUE #( tdformat = '*' tdline = lv_chunk ) TO rt_lines.
        lv_offset = lv_offset + strlen( lv_chunk ).
      ENDWHILE.
    ENDLOOP.
  ENDMETHOD.

  METHOD compare_text.
    rv_equal = abap_false.

    IF lines( it_expected ) <> lines( it_actual ).
      RETURN.
    ENDIF.

    LOOP AT it_expected INTO DATA(ls_expected).
      READ TABLE it_actual INDEX sy-tabix INTO DATA(ls_actual).

      IF sy-subrc <> 0
         OR ls_expected-tdformat <> ls_actual-tdformat
         OR ls_expected-tdline <> ls_actual-tdline.
        RETURN.
      ENDIF.
    ENDLOOP.

    rv_equal = abap_true.
  ENDMETHOD.

  METHOD add_log.
    DATA:
      ls_log   TYPE ty_log,
      ls_color TYPE lvc_s_scol.

    ls_log = VALUE #(
      row     = iv_row
      name    = iv_name
      status  = iv_status
      message = iv_message ).

    CASE iv_status.
      WHEN mc_status-error.
        ls_color = VALUE #(
          fname = 'STATUS'
          color = VALUE #( col = 6 int = 1 inv = 0 ) ).
      WHEN mc_status-updated.
        ls_color = VALUE #(
          fname = 'STATUS'
          color = VALUE #( col = 5 int = 1 inv = 0 ) ).
      WHEN mc_status-check.
        ls_color = VALUE #(
          fname = 'STATUS'
          color = VALUE #( col = 5 int = 0 inv = 0 ) ).
    ENDCASE.

    IF ls_color-fname IS NOT INITIAL.
      APPEND ls_color TO ls_log-color.
    ENDIF.

    APPEND ls_log TO mt_log.
    update_statistics( iv_status ).
  ENDMETHOD.

  METHOD update_statistics.
    CASE iv_status.
      WHEN mc_status-updated.
        ADD 1 TO mv_updated.
      WHEN mc_status-check.
        ADD 1 TO mv_checked.
      WHEN mc_status-error.
        ADD 1 TO mv_errors.
    ENDCASE.
  ENDMETHOD.

  METHOD set_column_texts.
    DATA lo_column TYPE REF TO cl_salv_column_table.

    TRY.
        lo_column ?= io_columns->get_column( 'ROW' ).
        lo_column->set_long_text( 'Строка Excel' ).
        lo_column->set_medium_text( 'Строка' ).
        lo_column->set_short_text( '№' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= io_columns->get_column( 'NAME' ).
        lo_column->set_long_text( 'Имя SO10 текста' ).
        lo_column->set_medium_text( 'Имя текста' ).
        lo_column->set_short_text( 'Имя' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= io_columns->get_column( 'STATUS' ).
        lo_column->set_long_text( 'Статус обработки' ).
        lo_column->set_medium_text( 'Статус' ).
        lo_column->set_short_text( 'Статус' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= io_columns->get_column( 'MESSAGE' ).
        lo_column->set_long_text( 'Сообщение' ).
        lo_column->set_medium_text( 'Сообщение' ).
        lo_column->set_short_text( 'Сообщение' ).
      CATCH cx_salv_not_found.
    ENDTRY.

    TRY.
        lo_column ?= io_columns->get_column( 'COLOR' ).
        lo_column->set_technical( abap_true ).
      CATCH cx_salv_not_found.
    ENDTRY.
  ENDMETHOD.

  METHOD set_top_of_list.
    DATA:
      lo_form  TYPE REF TO cl_salv_form_layout_grid,
      lo_label TYPE REF TO cl_salv_form_label,
      lo_text  TYPE REF TO cl_salv_form_text.

    lo_form = NEW cl_salv_form_layout_grid( ).

    lo_label = lo_form->create_label( row = 1 column = 1 ).
    lo_label->set_text( 'Результат загрузки SO10' ).

    lo_label = lo_form->create_label( row = 2 column = 1 ).
    lo_label->set_text( 'Всего строк:' ).

    lo_text = lo_form->create_text( row = 2 column = 2 ).
    lo_text->set_text( CONV string( mv_total ) ).

    lo_label = lo_form->create_label( row = 3 column = 1 ).
    lo_label->set_text( 'Успешно изменено:' ).

    lo_text = lo_form->create_text( row = 3 column = 2 ).
    lo_text->set_text( CONV string( mv_updated ) ).

    lo_label = lo_form->create_label( row = 4 column = 1 ).
    lo_label->set_text( 'Только проверено:' ).

    lo_text = lo_form->create_text( row = 4 column = 2 ).
    lo_text->set_text( CONV string( mv_checked ) ).

    lo_label = lo_form->create_label( row = 5 column = 1 ).
    lo_label->set_text( 'Ошибок:' ).

    lo_text = lo_form->create_text( row = 5 column = 2 ).
    lo_text->set_text( CONV string( mv_errors ) ).

    lo_label = lo_form->create_label( row = 6 column = 1 ).
    lo_label->set_text( 'Режим:' ).

    lo_text = lo_form->create_text( row = 6 column = 2 ).
    lo_text->set_text(
      COND string(
        WHEN p_test = abap_true
        THEN 'Только проверка'
        ELSE 'Изменение SO10' ) ).

    io_alv->set_top_of_list( lo_form ).
  ENDMETHOD.

  METHOD display_log.
    DATA lo_alv TYPE REF TO cl_salv_table.

    IF mt_log IS INITIAL.
      MESSAGE 'Обработка завершена. Лог пуст.' TYPE 'I'.
      RETURN.
    ENDIF.

    TRY.
        cl_salv_table=>factory(
          IMPORTING r_salv_table = lo_alv
          CHANGING  t_table      = mt_log ).

        lo_alv->get_functions( )->set_all( abap_true ).
        lo_alv->get_columns( )->set_optimize( abap_true ).
        lo_alv->get_columns( )->set_color_column( 'COLOR' ).

        set_column_texts( lo_alv->get_columns( ) ).
        set_top_of_list( lo_alv ).

        lo_alv->display( ).
      CATCH cx_salv_msg INTO DATA(lx_salv).
        MESSAGE lx_salv->get_text( ) TYPE 'I'.
    ENDTRY.
  ENDMETHOD.

  METHOD download_template.

    TYPES:
      BEGIN OF ty_template,
        name    TYPE string,
        content TYPE string,
      END OF ty_template.

    DATA:
      lt_data     TYPE STANDARD TABLE OF ty_template,
      lt_columns  TYPE if_fdt_doc_spreadsheet=>t_column,
      lv_xstring  TYPE xstring,
      lv_filename TYPE string,
      lv_path     TYPE string,
      lv_fullpath TYPE string,
      lv_action   TYPE i,
      lt_binary   TYPE solix_tab.

    DATA(lo_struct) = CAST cl_abap_structdescr(
      cl_abap_typedescr=>describe_by_data( VALUE ty_template( ) ) ).

    LOOP AT lo_struct->get_components( ) INTO DATA(ls_component).
      APPEND VALUE #(
        id           = sy-tabix
        name         = ls_component-name
        display_name = COND string(
          WHEN sy-tabix = 1 THEN 'Имя'
          WHEN sy-tabix = 2 THEN 'Содержимое'
          ELSE ls_component-name )
        is_result    = abap_true
        type         = ls_component-type ) TO lt_columns.
    ENDLOOP.

    TRY.
        lv_xstring = cl_fdt_xl_spreadsheet=>if_fdt_doc_spreadsheet~create_document(
          name          = 'SO10_TEMPLATE.xlsx'
          itab          = REF #( lt_data )
          columns       = lt_columns
          iv_call_type  = if_fdt_doc_spreadsheet=>gc_call_dec_table
          iv_sheet_name = 'SO10' ).

      CATCH cx_fdt_excel_core INTO DATA(lx_excel).
        MESSAGE lx_excel->get_text( ) TYPE 'E'.
    ENDTRY.

    IF lv_xstring IS INITIAL.
      MESSAGE 'Не удалось сформировать Excel-шаблон' TYPE 'E'.
    ENDIF.

    cl_gui_frontend_services=>file_save_dialog(
      EXPORTING
        window_title      = 'Сохранить шаблон Excel'
        default_extension = 'xlsx'
        default_file_name = 'SO10_TEMPLATE.xlsx'
        file_filter       = 'Excel (*.xlsx)|*.xlsx|'
      CHANGING
        filename          = lv_filename
        path              = lv_path
        fullpath          = lv_fullpath
        user_action       = lv_action
      EXCEPTIONS
        OTHERS = 1 ).

    IF sy-subrc <> 0 OR lv_action <> cl_gui_frontend_services=>action_ok.
      RETURN.
    ENDIF.

    lt_binary = cl_bcs_convert=>xstring_to_solix(
      iv_xstring = lv_xstring ).

    cl_gui_frontend_services=>gui_download(
      EXPORTING
        filename     = lv_fullpath
        filetype     = 'BIN'
        bin_filesize = xstrlen( lv_xstring )
      CHANGING
        data_tab     = lt_binary
      EXCEPTIONS
        OTHERS = 1 ).

    IF sy-subrc <> 0.
      MESSAGE 'Ошибка сохранения Excel-шаблона' TYPE 'E'.
    ENDIF.

    MESSAGE 'Шаблон Excel успешно сохранён' TYPE 'S'.
  ENDMETHOD.

ENDCLASS.

INITIALIZATION.
  btn_tpl = |Download template|.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  NEW lcl_application( )->f4_file( ).

AT SELECTION-SCREEN.
  IF sy-ucomm = 'TPL'.
    NEW lcl_application( )->download_template( ).
  ENDIF.

START-OF-SELECTION.
  NEW lcl_application( )->run( ).
