<?php

/**
 * @param $form
 * @param $form_state
 *
 * @return mixed
 */
function f4u_excel_inkasso_orders_export_form($form, &$form_state) {
  $form = [];

  $form['filters']['startDate'] = [
    '#type' => 'date_popup',
    '#date_format' => 'd-m-Y',
    '#title' => t('Start date:'),
    '#date_label_position' => 'none',
  ];
  $form['filters']['endDate'] = [
    '#type' => 'date_popup',
    '#date_format' => 'd-m-Y',
    '#title' => t('End date:'),
    '#date_label_position' => 'none',
  ];
  $form['actions']['submit'] = [
    '#type' => 'submit',
    '#value' => t('Export invoices'),
  ];

  $custom_form = system_settings_form($form);

  $custom_form['actions']['submit'] = [
    '#type' => 'submit',
    '#value' => t('Export invoice'),
  ];

  $custom_form['#submit'][] = 'f4u_excel_inkasso_orders_export_form_submit';
  return $custom_form;
}

/**
 * @param $form
 * @param $form_state
 */
function f4u_excel_inkasso_orders_export_form_validate($form, &$form_state) {
  if (strtotime($form_state['input']['startDate']['date']) > strtotime($form_state['input']['endDate']['date'])) {
    form_set_error('time', t('End  must be greater than the start date.'));
  }
}

/**
 * Callback for hook_menu.
 *
 * @return string
 * @throws \PhpOffice\PhpSpreadsheet\Exception
 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
 */
function f4u_excel_inkasso_orders_export_form_submit($form, &$form_state) {

  require_once f4ucoupon_get_library('PhpSpreadsheet') . '/src/Bootstrap.php';

  $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
  $spreadsheet->setActiveSheetIndex(0);
  $spreadsheet->getActiveSheet()->setCellValue('A1', 'Amount');
  $spreadsheet->getActiveSheet()->setCellValue('B1', 'All Amount');
  $spreadsheet->getActiveSheet()->setCellValue('C1', '№ Invoice');
  $spreadsheet->getActiveSheet()->setCellValue('D1', 'Product');
  $spreadsheet->getActiveSheet()->setCellValue('E1', 'Order date');
  $spreadsheet->getActiveSheet()->setCellValue('F1', 'Last date of reminder');
  $spreadsheet->getActiveSheet()->setCellValue('G1', 'Form of address');
  $spreadsheet->getActiveSheet()->setCellValue('H1', 'Last name');
  $spreadsheet->getActiveSheet()->setCellValue('I1', 'First name');
  $spreadsheet->getActiveSheet()->setCellValue('J1', 'Company name');
  $spreadsheet->getActiveSheet()->setCellValue('K1', 'Street + no');
  $spreadsheet->getActiveSheet()->setCellValue('L1', 'Country');
  $spreadsheet->getActiveSheet()->setCellValue('M1', 'Zip code');
  $spreadsheet->getActiveSheet()->setCellValue('N1', 'Location');
  $spreadsheet->getActiveSheet()->setCellValue('O1', 'Phone');
  $spreadsheet->getActiveSheet()->setCellValue('P1', 'Email');

  $dates = [
    strtotime(date('Y-m-d', strtotime($form_state['input']['startDate']['date']))),
    strtotime(date('Y-m-d', strtotime($form_state['input']['endDate']['date'] . '+ 1 day'))),
  ];

  $query = db_select('commerce_order', 'co');
  $query->innerJoin('commerce_invoice', 'ci', 'ci.order_id = co.order_id');
  $query->fields('co', ['order_id']);
  $query->condition('co.created', $dates, 'BETWEEN');
  $query->condition('co.status', 'collector_agency', '=');
  $invoices = $query->execute();

  foreach ($invoices->fetchAll() as $key => $value) {

    $order = entity_metadata_wrapper('commerce_order', $value->order_id);
    $invoice = commerce_invoice_load_by_order_id($value->order_id);
    $line_item = $order->commerce_line_items->value();
    $sku = '';
    $shipping = 0;

    foreach ($line_item as $item) {
      if ($item->type == 'event_coupon') {
        $product = commerce_product_load($item->commerce_product['und'][0]['product_id']);
        $sku = $item->line_item_label . '(' . $product->field_event_location['und'][0]['city'] . ')';
      }

      if ($item->type == 'shipping' || $item->type == 'express_delivery') {
        $shipping += $item->commerce_unit_price['und'][0]['amount'];
      }
    }

    $all_amount = commerce_currency_format(($order->commerce_order_total->value()['amount'] - $shipping), 'EUR');
    $amount = commerce_currency_format($order->commerce_order_total->value()['amount'], 'EUR');
    $invoice_number = '';

    if (isset($invoice->invoice_number) && $invoice->invoice_number !== NULL) {
      $invoice_number = $invoice->invoice_number;
    }

    $last_reminder_date = '';
    if (date('d.m.Y', $order->created->value()) != date('d.m.Y',
        strtotime(f4u_payment_reminder_get_current_due_days($order->getIdentifier()) . ' - 7 days'))) {
      $last_reminder_date = date('d.m.Y', strtotime(f4u_payment_reminder_get_current_due_days($order->getIdentifier()) . ' - 7 days'));
    }

    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(1, $key + 2, $amount);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(2, $key + 2, _f4u_payment_reminder_get_previous_fee($order->order_id->value(), 5));
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(3, $key + 2, $invoice_number);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(4, $key + 2, $sku);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(5, $key + 2, date('d.m.y', $order->created->value()));
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(6, $key + 2, $last_reminder_date);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(7, $key + 2, $order->commerce_customer_shipping->field_name_prefix_shipping->value()->name);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(8, $key + 2, $order->commerce_customer_billing->commerce_customer_address->value()['last_name']);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(9, $key + 2, $order->commerce_customer_billing->commerce_customer_address->value()['first_name']);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(10, $key + 2, $order->commerce_customer_billing->commerce_customer_address->value()['organisation_name']);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(11, $key + 2, $order->commerce_customer_billing->commerce_customer_address->value()['thoroughfare']);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(12, $key + 2, $order->commerce_customer_billing->commerce_customer_address->value()['country']);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(13, $key + 2, $order->commerce_customer_billing->commerce_customer_address->value()['postal_code']);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(14, $key + 2, $order->commerce_customer_billing->commerce_customer_address->value()['locality']);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(15, $key + 2, $order->commerce_customer_billing->commerce_customer_address->value()['mobile_number']);
    $spreadsheet->getActiveSheet()
      ->setCellValueByColumnAndRow(16, $key + 2, $order->mail->value());
  }

  foreach (range('A', 'R') as $columnID) {
    $spreadsheet->getActiveSheet()->getColumnDimension($columnID)
      ->setAutoSize(TRUE);
  }

  $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xls");

  $excel_uri = 'public://excel_invoices_inkasso/';
  $file_name = 'invoice_' . date('y_m_d', time()) . '.xls';

  if (!drupal_realpath($excel_uri)) {
    $writer->save($excel_uri . $file_name);
  }
  else {
    drupal_mkdir($excel_uri);
    $writer->save($excel_uri . $file_name);
  }

  header("Content-type: application/xls");
  header("Content-disposition: attachment; filename=" . $file_name);
  header("Content-type: application/xls");
  readfile($excel_uri . $file_name);

  exit;
}
