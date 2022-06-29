*&---------------------------------------------------------------------*
*& Report  ZABAP_TABL_UPLOAD
*&
*&---------------------------------------------------------------------*
*& Assigned By    : Aditya Jaieel PwC
*& Developed By   : Aditya Jaieel PwC
*& Developed On   : 09/01/2015
*& Request Number : BSDK900497
*& Task Number    : BSDK900498
*& T-Code         : ZCARVE46
*& FTS Document   :
*& Description    : Upload Data to table
*&---------------------------------------------------------------------*
*&                        Modification logs
*&---------------------------------------------------------------------*
*& Assigned By    :
*& Changed By     :
*& Date           :
*& Request Number :
*& Description    :
*&---------------------------------------------------------------------*

REPORT zabap_tabl_upload.

TYPES: BEGIN OF ty_dd03l,
  tabname TYPE tabname,
  fieldname TYPE fieldname,
  keyflag TYPE keyflag,
  mandatory TYPE mandatory,
  END OF ty_dd03l,

  BEGIN OF ty_excel_fields,
  index TYPE numc4,
  fieldname TYPE fieldname,
  END OF ty_excel_fields.

DATA: wa_dd02l TYPE dd02l,
      it_dd03l TYPE TABLE OF ty_dd03l,
      it_excel_fields TYPE TABLE OF ty_excel_fields,
      lv_fields_missing TYPE xfeld,
      lv_answer TYPE c.

FIELD-SYMBOLS: <fs_dd03l> TYPE ty_dd03l,
               <fs_excel_fields> TYPE ty_excel_fields.

* For excel worksheet logic
DATA: oref_container   TYPE REF TO cl_gui_custom_container,
      iref_control     TYPE REF TO i_oi_container_control,
      iref_document    TYPE REF TO i_oi_document_proxy,
      iref_spreadsheet TYPE REF TO i_oi_spreadsheet,
      iref_error       TYPE REF TO i_oi_error,
      v_document_url   TYPE        char256,
      i_sheets         TYPE        soi_sheets_table,
      i_data           TYPE        soi_generic_table,
      i_data_tmp       TYPE        soi_generic_table,
      i_ranges         TYPE        soi_range_list,
      lv_sheet_name    TYPE        soi_field_name,
      lv_rowoffset     TYPE        i,
      lv_columns       TYPE        i.

FIELD-SYMBOLS: <fs_sheets>     TYPE soi_sheets,
               <fs_data>       TYPE soi_generic_item.

* Constants
DATA: lc_grant_upl TYPE viewgrant VALUE space,
      lc_transp TYPE tabclass VALUE 'TRANSP',
      lc_cols TYPE i VALUE 1000.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE text-001.
PARAMETERS: p_tabl TYPE dd02l-tabname OBLIGATORY,
            p_file TYPE rlgrap-filename,
            p_rtn TYPE xfeld AS CHECKBOX.
SELECTION-SCREEN SKIP.
SELECTION-SCREEN COMMENT /1(60) c_1.
SELECTION-SCREEN COMMENT /1(60) c_2.
SELECTION-SCREEN END OF BLOCK b1.

INITIALIZATION.
  c_1 = text-013.
  c_2 = text-014.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  CALL FUNCTION 'F4_FILENAME'
    EXPORTING
      program_name  = syst-cprog
      dynpro_number = syst-dynnr
    IMPORTING
      file_name     = p_file.

AT SELECTION-SCREEN.
  IF sy-ucomm EQ 'ONLI'.
    PERFORM: check_tab,
             check_xls.
  ENDIF.

START-OF-SELECTION.
  lv_answer = '1'.
  IF lv_fields_missing EQ 'X' AND p_rtn <> 'X'.
    CALL FUNCTION 'POPUP_TO_CONFIRM'
      EXPORTING
        titlebar              = text-015
        text_question         = text-016
        display_cancel_button = ' '
      IMPORTING
        answer                = lv_answer
      EXCEPTIONS
        text_not_found        = 1
        OTHERS                = 2.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
       WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
  ENDIF.

  IF lv_answer EQ '1'.
    PERFORM: upload_excel.
  ELSE.
    MESSAGE text-017 TYPE 'S' DISPLAY LIKE 'E'.
    STOP.
  ENDIF.
*&---------------------------------------------------------------------*
*&      Form  check_tab
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM check_tab.
  CLEAR: wa_dd02l, it_dd03l[].
  SELECT SINGLE *
    FROM dd02l
    INTO wa_dd02l
    WHERE tabname EQ p_tabl.
  IF sy-subrc EQ 0.
    IF wa_dd02l-viewgrant <> lc_grant_upl.
      MESSAGE text-002 TYPE 'S' DISPLAY LIKE 'E'.
      STOP.
    ENDIF.

    IF wa_dd02l-tabclass <> lc_transp.
      MESSAGE text-003 TYPE 'S' DISPLAY LIKE 'E'.
      STOP.
    ENDIF.
  ELSE.
    CLEAR: wa_dd02l.
  ENDIF.

  IF wa_dd02l-tabname IS INITIAL.
    MESSAGE text-004 TYPE 'E'.
  ENDIF.

  SELECT tabname fieldname keyflag mandatory
    FROM dd03l
    INTO TABLE it_dd03l
    WHERE tabname EQ wa_dd02l-tabname
      AND precfield EQ space.
ENDFORM.                    "check_tab

*&---------------------------------------------------------------------*
*&      Form  check_xls
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM check_xls.
  DATA: lv_fields TYPE int4,
        lv_msg TYPE string.
  CLEAR: it_excel_fields[], lv_fields_missing.

  DESCRIBE TABLE it_dd03l LINES lv_fields.

  PERFORM: read_excel_header,
           read_excel_data USING lv_sheet_name 1 1 1 lv_fields.

  LOOP AT it_dd03l ASSIGNING <fs_dd03l> WHERE fieldname <> 'MANDT'.
    READ TABLE i_data ASSIGNING <fs_data> WITH KEY
    value = <fs_dd03l>-fieldname.
    IF sy-subrc <> 0.
      lv_fields_missing = 'X'.
      IF <fs_dd03l>-keyflag IS INITIAL AND <fs_dd03l>-mandatory IS INITIAL.
        lv_msg = text-005.
        REPLACE ALL OCCURRENCES OF '&' IN lv_msg WITH <fs_dd03l>-fieldname.
        MESSAGE lv_msg TYPE 'W'.
      ELSE.
        lv_msg = text-006.
        REPLACE ALL OCCURRENCES OF '&' IN lv_msg WITH <fs_dd03l>-fieldname.
        MESSAGE lv_msg TYPE 'E'.
      ENDIF.
    ELSE.
      APPEND INITIAL LINE TO it_excel_fields ASSIGNING <fs_excel_fields>.
      <fs_excel_fields>-index = <fs_data>-column.
      <fs_excel_fields>-fieldname = <fs_dd03l>-fieldname.
    ENDIF.
  ENDLOOP.
  SORT it_excel_fields BY index.
  DELETE ADJACENT DUPLICATES FROM it_excel_fields COMPARING index.

  IF it_excel_fields[] IS INITIAL.
    MESSAGE text-008 TYPE 'S' DISPLAY LIKE 'E'.
    STOP.
  ENDIF.

  "Check whether any data exists
  PERFORM read_excel_data USING lv_sheet_name 2 1 1 lv_fields.
  IF i_data[] IS INITIAL.
    MESSAGE text-007 TYPE 'S' DISPLAY LIKE 'E'.
    STOP.
  ENDIF.
ENDFORM.                    "check_xls

*&---------------------------------------------------------------------*
*&      Form  open_workbook
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM open_workbook.
  IF iref_control IS INITIAL.
    c_oi_container_control_creator=>get_container_control(
      IMPORTING
        control = iref_control    " Container Control
        error   = iref_error    " Error Object
*        retcode =     " Error Code (Obsolete)
    ).
    PERFORM check_iref_error.

    CREATE OBJECT oref_container
      EXPORTING
        container_name              = 'CONT'
      EXCEPTIONS
        cntl_error                  = 1
        cntl_system_error           = 2
        create_error                = 3
        lifetime_error              = 4
        lifetime_dynpro_dynpro_link = 5
        OTHERS                      = 6.
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    iref_control->init_control(
      EXPORTING
        inplace_enabled          = 'X'
        r3_application_name      =  'EXCEL CONTAINER'   " Application Name
        parent                   =  oref_container   " Parent Container
      IMPORTING
        error                    =  iref_error   " Error Object
      EXCEPTIONS
        javabeannotsupported     = 1
        OTHERS                   = 2
    ).
    PERFORM check_iref_error.

    iref_control->get_document_proxy(
      EXPORTING
        document_type      =  soi_doctype_excel_sheet
      IMPORTING
        document_proxy     = iref_document
        error              = iref_error
    ).
    PERFORM check_iref_error.
  ENDIF.

  CONCATENATE 'FILE://' p_file INTO v_document_url.

  iref_document->open_document(
    EXPORTING
      document_title   = 'Excel'
      document_url     = v_document_url
      open_inplace     = 'X'
    IMPORTING
      error            = iref_error
  ).
  PERFORM check_iref_error.

  iref_document->get_spreadsheet_interface(
    IMPORTING
      error           = iref_error    " Error?
      sheet_interface = iref_spreadsheet
  ).
  PERFORM check_iref_error.
ENDFORM.                    "open_workbook

*&---------------------------------------------------------------------*
*&      Form  close_workbook
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM close_workbook.
  IF NOT iref_document IS INITIAL.
    iref_document->close_document(
*      EXPORTING
*        do_save     = ' '
*        no_flush    = ' '
*      IMPORTING
*        error       =
*        has_changed =
*        retcode     =
    ).
    iref_document->release_document(
*      EXPORTING
*        no_flush = ' '
*      IMPORTING
*        error    =
*        retcode  =
    ).
  ENDIF.
ENDFORM.                    "close_workbook

*&---------------------------------------------------------------------*
*&      Form  READ_EXCEL_HEADER
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM read_excel_header.
  PERFORM open_workbook.
*  iref_spreadsheet->get_sheets(
*    EXPORTING
*      no_flush = ' '    " Flush?
*    IMPORTING
*      sheets   = i_sheets    " Names of Worksheets
*      error    = iref_error    " Error?
*  ).
*  PERFORM check_iref_error.

  iref_spreadsheet->get_active_sheet(
*    EXPORTING
*      no_flush  = ' '    " Flush?
      IMPORTING
        sheetname = lv_sheet_name    " Name of Worksheet
        error     = iref_error    " Error?
*      retcode   =     " text
    ).
  PERFORM check_iref_error.
  PERFORM close_workbook.
ENDFORM.                    "READ_EXCEL_HEADER

*&---------------------------------------------------------------------*
*&      Form  CHECK_IREF_ERROR
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM check_iref_error.
  IF iref_error->has_failed = 'X'.
    iref_error->raise_message(
      EXPORTING
        type           =     'E'
      EXCEPTIONS
        message_raised = 1
        flush_failed   = 2
        OTHERS         = 3
    ).
    IF sy-subrc <> 0.
      PERFORM close_workbook.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                 WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
  ENDIF.
ENDFORM.                    "CHECK_IREF_ERROR

*&---------------------------------------------------------------------*
*&      Form  read_excel_data
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->I_SHEET    text
*      -->I_STARTROW text
*      -->I_STARTCOL text
*      -->I_ROWS     text
*      -->I_COLUMNS  text
*----------------------------------------------------------------------*
FORM read_excel_data USING i_sheet TYPE soi_field_name
                           i_startrow TYPE i
                           i_startcol TYPE i
                           i_rows TYPE i
                           i_columns TYPE i.

  CLEAR: i_data[], i_data_tmp[].

  PERFORM open_workbook.
  iref_spreadsheet->select_sheet(
        EXPORTING
          name     = i_sheet    " Name of Worksheet
        IMPORTING
          error    =  iref_error   " Error?
      ).
  PERFORM check_iref_error.

  "Delete existing ranges
  iref_spreadsheet->get_ranges_names(
    IMPORTING
      error    = iref_error    " Error?
      ranges   = i_ranges    " Names of the Ranges
  ).
  PERFORM check_iref_error.

  iref_spreadsheet->delete_ranges(
    EXPORTING
      ranges   = i_ranges    " List of Ranges
    IMPORTING
      error    =  iref_error   " Error?
  ).
  PERFORM check_iref_error.

  iref_spreadsheet->set_selection(
    EXPORTING
      left     =  i_startcol   " Line of Top Left-Hand Cell
      top      =  i_startrow   " Column of Top Left-Hand Cell
      rows     =  i_rows   " Lines
      columns  =  i_columns   " Columns
    IMPORTING
      error    = iref_error    " Error?
  ).
  PERFORM check_iref_error.

  iref_spreadsheet->insert_range(
    EXPORTING
      columns  = i_columns   " Number of Columns
      rows     = i_rows    " Number of Rows
      name     = 'Test'    " Name of the Range
    IMPORTING
      error    = iref_error    " Errors
  ).
  PERFORM check_iref_error.

  iref_spreadsheet->get_ranges_data(
    EXPORTING
      all       = 'X'    " Get All Ranges?
    IMPORTING
      contents  =  i_data       " Contents of the Tables
      error     =  iref_error   " Errors?
    CHANGING
      ranges    =  i_ranges   " Specified Ranges
  ).
  PERFORM check_iref_error.

  DELETE i_data WHERE value IS INITIAL.
  PERFORM close_workbook.
ENDFORM.                    "read_excel_data

*&---------------------------------------------------------------------*
*&      Form  upload_excel
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM upload_excel.
  DATA: lv_linetype TYPE REF TO cl_abap_structdescr,
        lv_tabletype TYPE REF TO cl_abap_tabledescr,
        lv_data TYPE REF TO data,
        lv_datal TYPE REF TO data,
        lv_startrow TYPE int4,
        lv_endrow TYPE int4,
        lv_lines TYPE int4,
        lv_cols TYPE int4,
        lv_updated TYPE int4.

  FIELD-SYMBOLS: <fs_line> TYPE any,
                 <fs_fldval> TYPE any,
                 <fs_table> TYPE STANDARD TABLE.

  TRY .
      lv_linetype ?= cl_abap_typedescr=>describe_by_name( p_name = p_tabl ).
      CREATE DATA lv_datal TYPE HANDLE lv_linetype.
      ASSIGN lv_datal->* TO <fs_line>.

      lv_tabletype = cl_abap_tabledescr=>create( p_line_type  = lv_linetype ).
      CREATE DATA lv_data TYPE HANDLE lv_tabletype.
      ASSIGN lv_data->* TO <fs_table>.

      DESCRIBE TABLE it_dd03l LINES lv_cols.
      lv_startrow = 2.

      DO .
        CLEAR: <fs_table>[], <fs_line>.
        PERFORM read_excel_data USING lv_sheet_name lv_startrow 1 lc_cols lv_cols.
        IF i_data[] IS INITIAL.
          EXIT.
        ENDIF.

        SORT i_data BY row column.

        LOOP AT i_data ASSIGNING <fs_data>.
          AT NEW row.
            CLEAR: <fs_line>.
          ENDAT.

          READ TABLE it_excel_fields ASSIGNING <fs_excel_fields> WITH KEY
          index = <fs_data>-column BINARY SEARCH.
          IF sy-subrc EQ 0.
            ASSIGN COMPONENT <fs_excel_fields>-fieldname OF STRUCTURE <fs_line>
            TO <fs_fldval>.
            IF sy-subrc EQ 0.
              <fs_fldval> = <fs_data>-value.
            ENDIF.
          ENDIF.

          AT END OF row.
            APPEND <fs_line> TO <fs_table>.
          ENDAT.
        ENDLOOP.

        IF <fs_table>[] IS NOT INITIAL.
          PERFORM: merge_data USING <fs_table>.
          DESCRIBE TABLE <fs_table> LINES lv_lines.
          MODIFY (p_tabl) FROM TABLE <fs_table>.
          IF sy-subrc EQ 0.
            lv_updated = sy-dbcnt.
            COMMIT WORK.
          ELSE.
            lv_updated = 0.
          ENDIF.
        ENDIF.

        lv_endrow = lv_startrow + lv_lines - 1. "As start row is also selected
        WRITE:/ text-009, ' ', lv_startrow, ' ',
                text-010, ' ', lv_endrow, ' ',
                text-011, ' ', lv_updated.

        lv_startrow = lv_endrow + 1.
      ENDDO.
    CATCH cx_root.
* Error message
  ENDTRY.

ENDFORM.                    "upload_excel

*&---------------------------------------------------------------------*
*&      Form  merge_data
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
FORM merge_data USING i_data TYPE STANDARD TABLE.
  DATA: lv_linetype TYPE REF TO cl_abap_structdescr,
        lv_tabletype TYPE REF TO cl_abap_tabledescr,
        lv_data TYPE REF TO data,
        lv_datal TYPE REF TO data,
        lv_where TYPE string,
        lv_fields TYPE string.

  FIELD-SYMBOLS: <fs_line1> TYPE any,
                 <fs_line2> TYPE any,
                 <fs_line3> TYPE any,
                 <fs_fldval> TYPE any,
                 <fs_table2> TYPE STANDARD TABLE,
                 <fs_table3> TYPE STANDARD TABLE.
  IF p_rtn <> 'X' OR i_data[] IS INITIAL.
    RETURN.
  ENDIF.

  TRY .
      lv_linetype ?= cl_abap_typedescr=>describe_by_name( p_name = p_tabl ).
      CREATE DATA lv_datal TYPE HANDLE lv_linetype.
      ASSIGN lv_datal->* TO <fs_line1>.
      CREATE DATA lv_datal TYPE HANDLE lv_linetype.
      ASSIGN lv_datal->* TO <fs_line2>.
      CREATE DATA lv_datal TYPE HANDLE lv_linetype.
      ASSIGN lv_datal->* TO <fs_line3>.

      lv_tabletype = cl_abap_tabledescr=>create( p_line_type  = lv_linetype ).
      CREATE DATA lv_data TYPE HANDLE lv_tabletype.
      ASSIGN lv_data->* TO <fs_table2>.
      CREATE DATA lv_data TYPE HANDLE lv_tabletype.
      ASSIGN lv_data->* TO <fs_table3>.

* Build where clause and transporting fields
      CLEAR: lv_where.
      LOOP AT it_dd03l ASSIGNING <fs_dd03l> WHERE keyflag EQ 'X' AND
        fieldname <> 'MANDT'.
        IF NOT lv_where IS INITIAL.
          CONCATENATE lv_where ' AND ' <fs_dd03l>-fieldname ' EQ i_data-'
          <fs_dd03l>-fieldname INTO lv_where.
        ELSE.
          CONCATENATE <fs_dd03l>-fieldname ' EQ i_data-'
          <fs_dd03l>-fieldname INTO lv_where.
        ENDIF.
      ENDLOOP.

      IF NOT lv_where IS INITIAL.
        SELECT *
          FROM (p_tabl)
          INTO TABLE <fs_table2>
          FOR ALL ENTRIES IN i_data
          WHERE (lv_where).
      ENDIF.

      IF NOT <fs_table2>[] IS INITIAL.
        CLEAR: lv_where, lv_fields.
        LOOP AT it_dd03l ASSIGNING <fs_dd03l> WHERE fieldname <> 'MANDT'.
          IF <fs_dd03l>-keyflag EQ 'X'.
            IF NOT lv_where IS INITIAL.
              CONCATENATE lv_where ' AND ' <fs_dd03l>-fieldname ' EQ <fs_line2>-'
              <fs_dd03l>-fieldname INTO lv_where.
            ELSE.
              CONCATENATE <fs_dd03l>-fieldname ' EQ <fs_line2>-'
              <fs_dd03l>-fieldname INTO lv_where.
            ENDIF.
          ENDIF.

          READ TABLE it_excel_fields ASSIGNING <fs_excel_fields> WITH KEY
          fieldname = <fs_dd03l>-fieldname.
          IF sy-subrc <> 0.
            CONCATENATE lv_fields <fs_dd03l>-fieldname INTO lv_fields
            SEPARATED BY space.
          ENDIF.
        ENDLOOP.
        LOOP AT <fs_table2> ASSIGNING <fs_line2>.
          MODIFY i_data FROM <fs_line2> TRANSPORTING (lv_fields) WHERE (lv_where).
        ENDLOOP.
      ENDIF.
    CATCH cx_root.
  ENDTRY.
ENDFORM.                    "merge_data
