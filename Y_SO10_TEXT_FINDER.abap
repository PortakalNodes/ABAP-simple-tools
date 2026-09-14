REPORT y_so10_text_finder.
*---------------------------------------------------------------------*
* SO10 Text Search
*
* Search in SAPscript standard texts (SO10)
*
*---------------------------------------------------------------------*
TABLES: stxh.
*---------------------------------------------------------------------*
* Selection screen
*---------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME.

PARAMETERS:
  p_search TYPE string LOWER CASE.  " Text to search

SELECT-OPTIONS:
  s_tdname FOR stxh-tdname NO INTERVALS,               " Text name pattern
  s_tdid   FOR stxh-tdid DEFAULT 'ST' NO INTERVALS.    " Text ID

PARAMETERS:
  p_case   AS CHECKBOX DEFAULT space,                  " Respect case?
  p_alllan AS CHECKBOX DEFAULT space USER-COMMAND lan, " Search in all languages
  p_langu  TYPE sylangu DEFAULT sy-langu.

SELECTION-SCREEN END OF BLOCK b1.

*---------------------------------------------------------------------*
* Class definition
*---------------------------------------------------------------------*
CLASS lcl_text_finder DEFINITION CREATE PRIVATE.

  PUBLIC SECTION.

    CLASS-METHODS:
      get_instance
        RETURNING VALUE(ro_instance) TYPE REF TO lcl_text_finder.

    METHODS:
      run,
      on_selection_screen,
      f4_language,
      f4_text_id IMPORTING iv_fieldname TYPE dynfnam,
      on_double_click
          FOR EVENT double_click OF cl_salv_events_table
        IMPORTING
          row
          column.

  PRIVATE SECTION.

    CLASS-DATA:
      go_instance TYPE REF TO lcl_text_finder.

    TYPES:
      BEGIN OF ty_text_header,
        tdobject TYPE stxh-tdobject,
        tdname   TYPE stxh-tdname,
        tdid     TYPE stxh-tdid,
        tdspras  TYPE stxh-tdspras,
      END OF ty_text_header,

      BEGIN OF ty_result,
        tdname   TYPE stxh-tdname,
        tdid     TYPE stxh-tdid,
        tdspras  TYPE stxh-tdspras,
        line_no  TYPE i,
        tdformat TYPE tline-tdformat,
        tdline   TYPE tline-tdline,
        tdobject TYPE stxh-tdobject,
      END OF ty_result,

      ty_result_tab TYPE STANDARD TABLE OF ty_result WITH EMPTY KEY.

    DATA:
      mo_salv   TYPE REF TO cl_salv_table,
      mt_result TYPE ty_result_tab,
      mv_search TYPE string.

    METHODS:
      validate_selection,
      search_so10,
      display_alv.

ENDCLASS.

*---------------------------------------------------------------------*
* Selection screen output
*---------------------------------------------------------------------*
AT SELECTION-SCREEN OUTPUT.
  lcl_text_finder=>get_instance( )->on_selection_screen( ).

*---------------------------------------------------------------------*
* F4 Language
*---------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_langu.
  lcl_text_finder=>get_instance( )->f4_language( ).

*---------------------------------------------------------------------*
* F4 Text ID - LOW
*---------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR s_tdid-low.
  lcl_text_finder=>get_instance( )->f4_text_id( 'S_TDID-LOW' ).

*---------------------------------------------------------------------*
* Start
*---------------------------------------------------------------------*
START-OF-SELECTION.
  lcl_text_finder=>get_instance( )->run( ).

*---------------------------------------------------------------------*
* Class implementation
*---------------------------------------------------------------------*
CLASS lcl_text_finder IMPLEMENTATION.

  METHOD get_instance.
    IF go_instance IS NOT BOUND.
      go_instance = NEW lcl_text_finder( ).
    ENDIF.
    ro_instance = go_instance.
  ENDMETHOD.

  METHOD run.

    validate_selection( ).
    search_so10( ).

    IF mt_result IS INITIAL.
      MESSAGE |No SO10 texts found for "{ p_search }"| TYPE 'S' DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    display_alv( ).

  ENDMETHOD.
*---------------------------------------------------------------------*
* Selection screen output (show/hide language field)
*---------------------------------------------------------------------*
  METHOD on_selection_screen.
    LOOP AT SCREEN.
      IF screen-name CS 'P_LANGU'.
        IF p_alllan IS INITIAL.
          screen-active = '1'.
        ELSE.
          screen-active = '0'.
        ENDIF.
        MODIFY SCREEN.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

*---------------------------------------------------------------------*
* Validate selection
*---------------------------------------------------------------------*
  METHOD validate_selection.

    IF condense( p_search ) = ''.
      MESSAGE |Search text cannot be empty| TYPE 'S' DISPLAY LIKE 'E'.
      LEAVE LIST-PROCESSING.
    ENDIF.

    IF   s_tdname[] IS INITIAL
     AND s_tdid[] IS INITIAL.
      MESSAGE |Specify Text Name or Text ID| TYPE 'E'.
    ENDIF.

  ENDMETHOD.

*---------------------------------------------------------------------*
* Search SO10
*---------------------------------------------------------------------*
  METHOD search_so10.
    DATA:
      lt_headers TYPE STANDARD TABLE OF ty_text_header WITH EMPTY KEY,
      lt_lines   TYPE STANDARD TABLE OF tline WITH EMPTY KEY,
      ls_thead   TYPE thead.

    CLEAR mt_result.
*---------------------------------------------------------------*
* Prepare search string
*--------------------------------------------------------------*
    mv_search = condense( p_search ).
*--------------------------------------------------------------*
* Read candidate headers from STXH
*--------------------------------------------------------------*
    IF p_alllan = abap_true.
      SELECT DISTINCT
        tdobject,
        tdname,
        tdid,
        tdspras
        FROM stxh
        INTO TABLE @lt_headers
        WHERE tdobject = 'TEXT'
          AND tdname   IN @s_tdname
          AND tdid     IN @s_tdid.

    ELSE.
      SELECT DISTINCT
        tdobject,
        tdname,
        tdid,
        tdspras
        FROM stxh
        INTO TABLE @lt_headers
        WHERE tdobject = 'TEXT'
          AND tdname   IN @s_tdname
          AND tdid     IN @s_tdid
          AND tdspras  = @p_langu.
    ENDIF.

    IF lt_headers IS INITIAL.
      RETURN.
    ENDIF.
*-----------------------------------------------------------------*
* Process candidate texts
*-----------------------------------------------------------------*
    LOOP AT lt_headers ASSIGNING FIELD-SYMBOL(<fs_header>).
      CLEAR:
        lt_lines,
        ls_thead.

      CALL FUNCTION 'READ_TEXT'
        EXPORTING
          id                      = <fs_header>-tdid
          language                = <fs_header>-tdspras
          name                    = <fs_header>-tdname
          object                  = <fs_header>-tdobject
        IMPORTING
          header                  = ls_thead
        TABLES
          lines                   = lt_lines
        EXCEPTIONS
          id                      = 1
          language                = 2
          name                    = 3
          not_found               = 4
          object                  = 5
          reference_check         = 6
          wrong_access_to_archive = 7
          OTHERS                  = 8.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.
*----------------------------------------------------------------*
* Search text lines
*----------------------------------------------------------------*
      LOOP AT lt_lines ASSIGNING FIELD-SYMBOL(<fs_line>).
        TRY.
            IF p_case IS NOT INITIAL.
              DATA(lv_result) = find( val = <fs_line>-tdline sub = mv_search case = abap_true ).
            ELSE.
              lv_result = find( val = <fs_line>-tdline sub = mv_search case = abap_false ).
            ENDIF.
          CATCH cx_sy_strg_par_val.
            CONTINUE.
        ENDTRY.

        CHECK lv_result >= 0.

        APPEND VALUE #(
          tdname   = <fs_header>-tdname
          tdid     = <fs_header>-tdid
          tdspras  = <fs_header>-tdspras
          tdobject = <fs_header>-tdobject
          line_no  = sy-tabix
          tdformat = <fs_line>-tdformat
          tdline   = <fs_line>-tdline ) TO mt_result.
      ENDLOOP.
    ENDLOOP.
*-----------------------------------------------------------------*
* Sort result
*-----------------------------------------------------------------*
    SORT mt_result BY
      tdname
      tdid
      tdspras
      line_no.

  ENDMETHOD.
*---------------------------------------------------------------------*
* Display SALV
*---------------------------------------------------------------------*
  METHOD display_alv.
    DATA:
      lo_columns TYPE REF TO cl_salv_columns_table,
      lo_column  TYPE REF TO cl_salv_column_table.

*-------------------------------------------------------------------*
* Create SALV
*-------------------------------------------------------------------*
    TRY.
        cl_salv_table=>factory(
          IMPORTING
            r_salv_table = mo_salv
          CHANGING
            t_table      = mt_result ).
      CATCH cx_salv_msg INTO DATA(lx_salv).
        MESSAGE lx_salv->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
        RETURN.
    ENDTRY.

*-------------------------------------------------------------------*
* Functions
*-------------------------------------------------------------------*
    mo_salv->get_functions( )->set_all( abap_true ).

*-------------------------------------------------------------------*
* Display settings
*-------------------------------------------------------------------*
    mo_salv->get_display_settings( )->set_striped_pattern( abap_true ).

    mo_salv->get_display_settings( )->set_list_header(
      |SO10 Search: "{ p_search }" - { lines( mt_result ) } matches| ).

*-------------------------------------------------------------------*
* Columns
*-------------------------------------------------------------------*
    lo_columns = mo_salv->get_columns( ).
    lo_columns->set_optimize( abap_true ).

    TRY.

        lo_column ?= lo_columns->get_column( 'TDNAME' ).
        lo_column->set_short_text( 'Text' ).
        lo_column->set_medium_text( 'SO10 Text' ).
        lo_column->set_long_text( 'SO10 Text Name' ).

        lo_column ?= lo_columns->get_column( 'TDID' ).
        lo_column->set_short_text( 'ID' ).
        lo_column->set_medium_text( 'Text ID' ).
        lo_column->set_long_text( 'SAPscript Text ID' ).

        lo_column ?= lo_columns->get_column( 'TDSPRAS' ).
        lo_column->set_short_text( 'Lang.' ).
        lo_column->set_medium_text( 'Language' ).
        lo_column->set_long_text( 'Text Language' ).

        lo_column ?= lo_columns->get_column( 'LINE_NO' ).
        lo_column->set_short_text( 'Line' ).
        lo_column->set_medium_text( 'Line' ).
        lo_column->set_long_text( 'Text Line' ).

        lo_column ?= lo_columns->get_column( 'TDFORMAT' ).
        lo_column->set_short_text( 'Format' ).
        lo_column->set_medium_text( 'Format' ).
        lo_column->set_long_text( 'SAPscript Format' ).

        lo_column ?= lo_columns->get_column( 'TDLINE' ).
        lo_column->set_short_text( 'Text' ).
        lo_column->set_medium_text( 'Found Text' ).
        lo_column->set_long_text( 'Found Text' ).

        lo_column ?= lo_columns->get_column( 'TDOBJECT' ).
        lo_column->set_visible( abap_false ).

      CATCH cx_salv_not_found.
        " Ignore column errors
    ENDTRY.
*-------------------------------------------------------------------*
* Sorting
*-------------------------------------------------------------------*
    TRY.
        DATA(lo_sorts) = mo_salv->get_sorts( ).

        lo_sorts->add_sort(
          columnname = 'TDNAME'
          sequence   = if_salv_c_sort=>sort_up ).

        lo_sorts->add_sort(
          columnname = 'TDID'
          sequence   = if_salv_c_sort=>sort_up ).

        lo_sorts->add_sort(
          columnname = 'TDSPRAS'
          sequence   = if_salv_c_sort=>sort_up ).

        lo_sorts->add_sort(
          columnname = 'LINE_NO'
          sequence   = if_salv_c_sort=>sort_up ).

      CATCH cx_salv_not_found cx_salv_existing cx_salv_data_error.
        " Ignore
    ENDTRY.
*-------------------------------------------------------------------*
* Register double click handler (this instance)
*-------------------------------------------------------------------*
    DATA(lo_events) = mo_salv->get_event( ).
    SET HANDLER on_double_click FOR lo_events.
*-------------------------------------------------------------------*
* Display
*-------------------------------------------------------------------*
    mo_salv->display( ).

  ENDMETHOD.

*---------------------------------------------------------------------*
* Double click -> open cached text in SAPscript editor (display mode)
*---------------------------------------------------------------------*
  METHOD on_double_click.
    DATA:
      lt_lines TYPE STANDARD TABLE OF tline WITH EMPTY KEY,
      ls_thead TYPE thead.

    READ TABLE mt_result ASSIGNING FIELD-SYMBOL(<ls_result>) INDEX row.
    IF sy-subrc <> 0.
      RETURN.
    ENDIF.

    CALL FUNCTION 'READ_TEXT'
      EXPORTING
        id                      = <ls_result>-tdid
        language                = <ls_result>-tdspras
        name                    = <ls_result>-tdname
        object                  = <ls_result>-tdobject
      IMPORTING
        header                  = ls_thead
      TABLES
        lines                   = lt_lines
      EXCEPTIONS
        id                      = 1
        language                = 2
        name                    = 3
        not_found               = 4
        object                  = 5
        reference_check         = 6
        wrong_access_to_archive = 7
        OTHERS                  = 8.
    IF sy-subrc <> 0.
      MESSAGE |READ_TEXT failed for "{ <ls_result>-tdname }", exception { sy-subrc }|
        TYPE 'S' DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

*-------------------------------------------------------------------*
* Open SAPscript editor in display mode
*-------------------------------------------------------------------*
    CALL FUNCTION 'EDIT_TEXT'
      EXPORTING
        display       = 'X'
        header        = ls_thead
        save          = space
        program       = sy-repid
      TABLES
        lines         = lt_lines
      EXCEPTIONS
        id            = 1
        language      = 2
        linesize      = 3
        name          = 4
        object        = 5
        textformat    = 6
        communication = 7
        OTHERS        = 8.
    IF sy-subrc <> 0.
      MESSAGE |EDIT_TEXT failed for "{ <ls_result>-tdname }", exception { sy-subrc }|
        TYPE 'S' DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

  ENDMETHOD.
*---------------------------------------------------------------------*
* F4 Language
*---------------------------------------------------------------------*
  METHOD f4_language.
    DATA lt_return TYPE STANDARD TABLE OF ddshretval.

    SELECT
      sprsl,
      sptxt
      FROM t002t
      INTO TABLE @DATA(lt_values)
      WHERE spras = @sy-langu.

    IF lt_values IS INITIAL.
      RETURN.
    ENDIF.

    SORT lt_values BY sptxt.

    CALL FUNCTION 'F4IF_INT_TABLE_VALUE_REQUEST'
      EXPORTING
        retfield    = 'SPRSL'
        dynpprog    = sy-repid
        dynpnr      = sy-dynnr
        dynprofield = 'P_LANGU'
        value_org   = 'S'
      TABLES
        value_tab   = lt_values
        return_tab  = lt_return
      EXCEPTIONS
        OTHERS      = 0.

  ENDMETHOD.
*---------------------------------------------------------------------*
* F4 Text ID
*---------------------------------------------------------------------*
  METHOD f4_text_id.
    DATA lt_return TYPE STANDARD TABLE OF ddshretval.

    SELECT
      t~tdobject,
      t~tdid,
      tdtext
      FROM ttxid AS t
      LEFT JOIN ttxit AS tt ON tt~tdobject = t~tdobject
                            AND tt~tdid    = t~tdid
      BYPASSING BUFFER
      INTO TABLE @DATA(lt_values)
      WHERE t~tdobject = 'TEXT'
        AND tt~tdspras = @sy-langu.

    IF lt_values IS INITIAL.
      RETURN.
    ENDIF.

    CALL FUNCTION 'F4IF_INT_TABLE_VALUE_REQUEST'
      EXPORTING
        retfield    = 'TDID'
        dynpprog    = sy-repid
        dynpnr      = sy-dynnr
        dynprofield = iv_fieldname
        value_org   = 'S'
      TABLES
        value_tab   = lt_values
        return_tab  = lt_return
      EXCEPTIONS
        OTHERS      = 0.

  ENDMETHOD.

ENDCLASS.
