<?php
include_once(__DIR__ . '/../../../../global_includes/excel/ExcelWriterXML.php');
include_once(__DIR__ . '/../../../classes/quote.php');
include_once(__DIR__ . '/../../../classes/invoice.php');
include_once(__DIR__ . '/../../../includes/date_functions.php');
include_once(__DIR__ . '/../../../includes/sorting_functions.php');
require_once(__DIR__ . '/../../../classes/payment_to_supplier.php');


$data = array();


$sort_by = (isset ($_GET['sort_by'])) ? $_GET['sort_by'] : 'invoice_number';
$search_field = filter_input(INPUT_GET, 'search_field');
$visible_paid_status = $_SESSION['table_params']['visible_paid_status'] = filter_input(INPUT_GET, 'visible_paid_status') ?: 'show';
$offset = 1;

$limit = false;

$data['oti_invoices'] = GetInvoices($search_field, $offset, $limit, $sort_by, $visible_paid_status);

//var_dump($data['oti_invoices']);

// PROCEED TO WRITE DOCUMENT
$xml = new ExcelWriterXML ('billsReport.xls');


$xml->docTitle($input['itinerary_name']);
$xml->docAuthor('_daemon');
$xml->docCompany('Overseas Travel International');
$xml->showErrorSheet(false);
$sheet = $xml->addSheet('Profit Report');

// returns an array of objects containing xml styles
$xml_styles = get_xml_styles($xml);


$agency_info = $input['agency_info'];
$financial = $input['financial'];
// Title


//
$sheet->WriteString(1, 1, "Vendor", $xml_styles['column_names']);
$sheet->WriteString(1, 2, "AP account", $xml_styles['column_names']);
$sheet->WriteString(1, 3, "Vendor address: City", $xml_styles['column_names']);
$sheet->WriteString(1, 4, "Vendor address: Country", $xml_styles['column_names']);
$sheet->WriteString(1, 5, "Vendor address: Line 1", $xml_styles['column_names']);
$sheet->WriteString(1, 6, "Vendor address: Line 2", $xml_styles['column_names']);
$sheet->WriteString(1, 7, "Vendor address: Line 3", $xml_styles['column_names']);
$sheet->WriteString(1, 8, "Vendor address: Line 4", $xml_styles['column_names']);
$sheet->WriteString(1, 9, "Vendor address: Line 5", $xml_styles['column_names']);
$sheet->WriteString(1, 10, "Vendor address: Note", $xml_styles['column_names']);
$sheet->WriteString(1, 11, "Vendor address: State", $xml_styles['column_names']);
$sheet->WriteString(1, 12, "Vendor address: Postal Code", $xml_styles['column_names']);
$sheet->WriteString(1, 13, "Due date", $xml_styles['column_names']);
$sheet->WriteString(1, 14, "Bill number", $xml_styles['column_names']);
$sheet->WriteString(1, 15, "Currency", $xml_styles['column_names']);
$sheet->WriteString(1, 16, "Discount amount", $xml_styles['column_names']);
$sheet->WriteString(1, 17, "Discount description", $xml_styles['column_names']);
$sheet->WriteString(1, 18, "Discount item", $xml_styles['column_names']);
$sheet->WriteString(1, 19, "Discount tax code", $xml_styles['column_names']);
$sheet->WriteString(1, 20, "Exchange rate", $xml_styles['column_names']);
//$sheet->WriteString(1, 21, "Expense Line: Account", $xml_styles['column_names']);
$sheet->WriteString(1, 21, "Expense Line: Amount", $xml_styles['column_names']);
$sheet->WriteString(1, 22, "Expense Line: Billable status", $xml_styles['column_names']);
$sheet->WriteString(1, 23, "Expense Line: Class", $xml_styles['column_names']);
$sheet->WriteString(1, 24, "Expense Line: Customer", $xml_styles['column_names']);
$sheet->WriteString(1, 25, "Expense Line: Memo", $xml_styles['column_names']);
$sheet->WriteString(1, 26, "Expense Line: Tax code", $xml_styles['column_names']);
$sheet->WriteString(1, 27, "Item Line: Amount", $xml_styles['column_names']);
$sheet->WriteString(1, 28, "Item Line: Billable status", $xml_styles['column_names']);
$sheet->WriteString(1, 29, "Item Line: Class", $xml_styles['column_names']);
$sheet->WriteString(1, 30, "Item Line: Cost", $xml_styles['column_names']);
$sheet->WriteString(1, 31, "Item Line: Customer", $xml_styles['column_names']);
$sheet->WriteString(1, 32, "Item Line: Description", $xml_styles['column_names']);
$sheet->WriteString(1, 33, "Item Line: Inventory bin", $xml_styles['column_names']);
$sheet->WriteString(1, 34, "Item Line: Inventory site", $xml_styles['column_names']);
$sheet->WriteString(1, 35, "Item Line: Item", $xml_styles['column_names']);
$sheet->WriteString(1, 36, "Item Line: Linked PO number", $xml_styles['column_names']);
$sheet->WriteString(1, 37, "Item Line: Quantity", $xml_styles['column_names']);
$sheet->WriteString(1, 38, "Item Line: Serial/Lot number", $xml_styles['column_names']);
$sheet->WriteString(1, 39, "Item Line: Tax code", $xml_styles['column_names']);
$sheet->WriteString(1, 40, "Item Line: U/M", $xml_styles['column_names']);
$sheet->WriteString(1, 41, "Linked PO number", $xml_styles['column_names']);
$sheet->WriteString(1, 42, "Memo", $xml_styles['column_names']);
$sheet->WriteString(1, 43, "Shipping account", $xml_styles['column_names']);
$sheet->WriteString(1, 44, "Shipping description", $xml_styles['column_names']);
$sheet->WriteString(1, 45, "Shipping item", $xml_styles['column_names']);
$sheet->WriteString(1, 46, "Shipping tax code", $xml_styles['column_names']);
$sheet->WriteString(1, 47, "Terms", $xml_styles['column_names']);
$sheet->WriteString(1, 48, "Transaction date", $xml_styles['column_names']);
$sheet->WriteString(1, 49, "Update Bill", $xml_styles['column_names']);


$row = 1; // row counter start
$dataArray = array();


foreach ($data['oti_invoices'] as $invoice_id => $invoice) {
//    var_dump($invoice['items']);
    // GET ALL RELATED ITEMS

    $invDate = $invoice['creation_date'];


    $invoice_id = $invoice_id;
    $invoiceNumber = $invoice['invoice_number'];

    $invoice = new Invoice ($invoice_id);


    $agency_info = $invoice->GetAgencyInfo();

    // INSTANTIATE QUOTE ITEM
    $quote = new Quote ($invoice->data->quote_id);
    $quote_data = $quote->GetQuote();

    // SET ALL EXPECTED VARIABLES FOR QUOTE
    $input['itinerary_id'] = $quote->GetQuote()['quote_id'];
    $input['itinerary_name'] = $quote->GetQuote()['quote_title'];

    $input['financial'] = $quote->GetQuote()['quote_title'];
    $input['itinerary_start'] = '';
    $input['itinerary_end'] = '';

    // ADD AGENCY INFORMATION TO OUR INPUT
    $input['agency_info'] = $agency_info;
    $input['agency_info']['agency'] = $agency_info['agency_name'];


    $itinerary_start = date("d-M-y", strtotime($input['itinerary_start']));

    $itinerary_end = date("d-M-y", strtotime($input['itinerary_end']));


    // GET ALL RELATED ITEMS
    $item_ids = $invoice->GetItemIDs();

    // STRUCTURE ITEMS AND ADD THEM TO OUR INPUT
    $itinerarxy_name = str_replace('(', ':', $input['itinerary_name']);
    $itinerarxy_name = str_replace(')', '', $itinerarxy_name);

    $arrayDescriptions = array();

    foreach ($item_ids as $key => $item_id) {
        //Vendor	AP account	Vendor address: City	Vendor address: Country	Vendor address: Line 1
        $dataItem = array(
            'vendor' => 'vendor',//TBD Vendor
            'vendorAddressCity' => 'city',//TBD
            'vendorAddressCountry' => 'country',//TBD
            'vendorAddressLine1' => 'address',//TBD
            'vendorAddressState' => 'state',//TBD
            'vendorAddressPostalCode' => 'postal code',//TBD
            'dueDate' => 'due date',//TBD CHECKING DATE Y-m-d h:j:s
            'billNumber' => 'bill',//TBD invoice number plus dash
            'expenseLineAccount' => 'Cost of goods sold',//Concat Department and office
            'itemLineBillableStatus' => 'HasBeenBilled',// HasBeenBilled
            'itemLineCost' => 'item line cost',//TBD total price
            'itemLineCustomer' => 'item line customer',//TBD Agency name plus customer name
            'itemLineDescription' => 'item line description',//TBD Checkin and checkout dates
            'itemLineItem' => 'item line item',//TBD hotel, activity, trasnfers...
            'itemLineQuantity' => '1',//TBD just 1
            'transactionDate' => 'transaction date',//TBD Invoice date
        );

        $item = new InvoiceItem($item_id);

        // Remove items not marked as visible in quote / ib
        $item_data = $item->reservation;

        $type = $item_data['reservation_type'] = $item->data->reservation_type;


        $item_data['modification_type'] = $item->data->modification_type;

        $parent_id = isset ($item_data['request']['parent_item_id']) ? $item_data['request']['parent_item_id'] : false;

        $itemTransactionData = $item->data;

        $transactionDate = $itemTransactionData->creation_date;

        //AGENT INFO
        $agentName = $item->data->created_by;


        $agentDept = $item->data->department;

        $agentOffice = '';

        $result = $dashboard_pdo->prepare("SELECT * FROM oti_users  WHERE  username = '$agentName'");
        if ($result->execute() && is_array($userList = $result->fetchAll())) {
            $agentOffice = $userList[0]->user_office;
        }
        if (!$agentOffice) {
            $agentOffice = 'Miami';
        }

        $items = $item_data;


        //RATES INFO
        $rates_data = $items['rates_data'];

        $daily_rates = $rates_data['daily_rates'];

        $units = $items['request']['units'];

        $currency = 'USD';

        //CLIENT CHECK IN AND OUT DATES
        $checkInDate = $items['request']['check_in_date'];
        $checkOutDate = $items['request']['check_out_date'];
        $duration = GetDuration($checkInDate, $checkOutDate);
        //      var_dump($duration);
        //      var_dump($checkInDate);
        //      var_dump($checkOutDate);
        //      var_dump($item_id);
        //      var_dump($item);
        //      var_dump($items);
        //      exit;

        //CLIENT INFO

        $clientInfo = $items['guest_info'][0];
        $clientLastName = $clientInfo['guest_last_name'];
        $clientFirstName = $clientInfo['guest_first_name'];


        for ($i = 0; $i < count($daily_rates); $i++) {
            $wrap_it_up = false; //INTERVAL END FLAG

            $rate_net = $daily_rates[$i]['rate_net'];

            $rate_retail_before_tax = $daily_rates[$i]['rate_retail_before_tax'];


            //SET HOTEL COMMISSION

            $rate_commission_percent_net = $daily_rates[$i]['rate_commission_percent_net'] / 100;

            //SET AGENCY COMMISSION

            $rate_commission_percent = isset ($rates_data['total']['rate_commission_percent']) ? $rates_data['total']['rate_commission_percent'] / 100 : 0;

            $current_date = $daily_rates[$i]['rate_date'];

            //  var_dump($current_date);
//            var_dump($interval_start);

            if (!isset ($interval_start)) {
                $interval_start = $current_date;
//                var_dump('----------------------');

            }

            //IF TODAY'S RATE_NET DOES NOT EQUAL TOMORROWS

            if ((!isset($daily_rates[$i + 1]) or $rate_net != $daily_rates[$i + 1]['rate_net'])) {
                $wrap_it_up = true;
            } //IF TODAYS RATE_COMISSION_NET DOES NOT EQUAL TOMORROWS

            else if ($daily_rates[$i]['rate_commission_percent_net'] != $daily_rates[$i + 1]['rate_commission_percent_net']) {
                $wrap_it_up = true;
            } // IF TODAY'S RATE_RETAIL DOES NOT EQUAL TOMORROWS
            else if ($daily_rates[$i]['rate_retail_before_tax'] != $daily_rates[$i + 1]['rate_retail_before_tax']) {
                $wrap_it_up = true;
            } // IF TODAY'S RATE_FEE DOES NOT EQUAL TOMORROWS
            else if ($daily_rates[$i]['rate_fee'] != $daily_rates[$i + 1]['rate_fee']) {

                $wrap_it_up = true;
            } // IF TODAY'S RATE_FEE_NET DOES NOT EQUAL TOMORROWS
            else if ($daily_rates[$i]['rate_fee_net'] != $daily_rates[$i + 1]['rate_fee_net']) {
                $wrap_it_up = true;
            }
            $interval_end = $current_date;

            // var_dump($interval_end);

            if ($wrap_it_up) {

                $rate_tax_percent = 1 + ($daily_rates[$i]['rate_tax_percent'] / 100);

                $rate_tax_percent_net = 1 + ($rates_data['total']['rate_tax_percent_net'] / 100);

                $rate_fee = $daily_rates[$i]['rate_fee'];

                $rate_fee_net = $daily_rates[$i]['rate_fee_net'];

                $rate_net = $daily_rates[$i]['rate_net'];

                $rate_retail = $daily_rates[$i]['rate_retail_before_tax'];

                $rate_commission_percent = $daily_rates[$i]['rate_commission_percent'] / 100;
                $rate_commission_percent_net = $daily_rates[$i]['rate_commission_percent_net'] / 100;

                $start_stamp = date_create_from_format("Y-m-d", $interval_start);
                //  var_dump($start_stamp);
                $end_stamp = date_create_from_format("Y-m-d", $interval_end);
                //  $start_stamp = new DateTime ($interval_start);
                $interval = $start_stamp->diff($end_stamp);

                $days = ($interval->days) + 1;
                //  $start_date = date("m/d/Y", strtotime($checkInDate));

//                   var_dump($days);

                $end_date = $items['reservation_type'] == 'hotel' || $items['reservation_type'] == 'resort_fee' ? date("m/d/Y", strtotime($interval_end) + 86400) : date("m/d/Y", strtotime($interval_end));

                $vendorAddressCountry = $vendorAddressCity = $vendorAddressLine1 = $vendorAddressState = $vendorAddressPostalCode = '';
                //    Build the item name depending on the source
                switch ($items['reservation_type']) {
                    case 'hotel':
                        $description = 'Hotel - ' . $items['hotel_info']['hotel_name'] . ' - ' . $items['room_info']['room_name'] . ' - ' . $items['room_info']['room_beds'];
                        $vendorName = $items['hotel_info']['hotel_name'];
                        $vendorAddressCity = $items['hotel_info']['hotel_city'];
                        $vendorAddressCountry = $items['hotel_info']['hotel_country'];
                        $vendorAddressLine1 = $items['hotel_info']['hotel_address'];
                        $vendorAddressState = $items['hotel_info']['hotel_state'];
                        $vendorAddressPostalCode = $items['hotel_info']['hotel_zip_code'];

                        // If rates were obtained from another supplier (bonotel, hotelbeds etc) add this to the name
                        if ($items['rates_data']['rate_supplier'] != '') {

                        }
                        break;
                    case 'activity':
                        $description = 'Activity - ' . $items['supplier_info']['supplier_name'] . ' - ' . $items['activity_info']['activity_name'];
                        $vendorName = $items['supplier_info']['supplier_name'];
                        $vendorAddressCity = $items['supplier_info']['supplier_city'];
                        $vendorAddressCountry = $items['supplier_info']['supplier_country'];
                        $vendorAddressLine1 = $items['supplier_info']['supplier_address'];
                        $vendorAddressState = $items['supplier_info']['supplier_state'];
                        $vendorAddressPostalCode = $items['supplier_info']['supplier_zip_code'];

                        $end_date = date("m/d/Y", strtotime($items['request']['check_out_date']));
                        break;
                    case 'transfer':
                        $description = 'Transfer - ' . $items['supplier_info']['supplier_name'] . ' - ' . $items['pickup_info']['pickup_name'] . " to " . $items['dropoff_info']['dropoff_name'];
                        $vendorName = $items['supplier_info']['supplier_name'];
                        $vendorAddressCity = $items['supplier_info']['supplier_city'];
                        $vendorAddressCountry = $items['supplier_info']['supplier_country'];
                        $vendorAddressLine1 = $items['supplier_info']['supplier_address'];
                        $vendorAddressState = $items['supplier_info']['supplier_state'];
                        $vendorAddressPostalCode = $items['supplier_info']['supplier_zip_code'];
                        $end_date = date("m/d/Y", strtotime($items['request']['check_out_date']));
                        break;
                    case 'resort_fee':
                        $description = 'Hotel Resort Fees - ' . $items['hotel_info']['hotel_name'] . ' - RF';
                        $vendorName = $items['hotel_info']['hotel_name'];
                        $vendorAddressCity = $items['hotel_info']['hotel_city'];
                        $vendorAddressCountry = $items['hotel_info']['hotel_country'];
                        $vendorAddressLine1 = $items['hotel_info']['hotel_address'];
                        $vendorAddressState = $items['hotel_info']['hotel_state'];
                        $vendorAddressPostalCode = $items['hotel_info']['hotel_zip_code'];

                        break;
                    case 'hotel_service':
                        $description = 'Hotel Service - ' . $items['hotel_info']['hotel_name'] . ' - ' . $items['service_info']['service_name'];
                        $vendorName = $items['hotel_info']['hotel_name'];
                        $vendorAddressCity = $items['hotel_info']['hotel_city'];
                        $vendorAddressCountry = $items['hotel_info']['hotel_country'];
                        $vendorAddressLine1 = $items['hotel_info']['hotel_address'];
                        $vendorAddressState = $items['hotel_info']['hotel_state'];
                        $vendorAddressPostalCode = $items['hotel_info']['hotel_zip_code'];
                        break;
                    case 'car_rental':
                        $description = 'Car Rental - ' . $items['supplier_info']['supplier_name'] . ' - ' . $items['request']['vehicle_group'];
                        $vendorName = $items['supplier_info']['supplier_name'];
                        $vendorAddressCity = $items['supplier_info']['supplier_city'];
                        $vendorAddressCountry = $items['supplier_info']['supplier_country'];
                        $vendorAddressLine1 = $items['supplier_info']['supplier_address'];
                        $vendorAddressState = $items['supplier_info']['supplier_state'];
                        $vendorAddressPostalCode = $items['supplier_info']['supplier_zip_code'];
                        break;
                    case 'taxable_fee':
                        if (!isset ($items['taxable_fee_info']['taxable_fee_name']) || $items['taxable_fee_info']['taxable_fee_name'] == '') {
                            $items['taxable_fee_info']['taxable_fee_name'] = 'Fee';
                        }

                        $description = 'Activity Fees - ' . $items['activity_info']['activity_name'] . ' - ' . $items['taxable_fee_info']['taxable_fee_name'];
                        $vendorName = $items['activity_info']['activity_name'];
                        $vendorAddressCity = $items['activity_info']['activity_city'];
                        $vendorAddressCountry = $items['activity_info']['activity_state'];
                        $vendorAddressLine1 = $items['activity_info']['activity_address'];
                        $vendorAddressState = $items['activity_info']['activity_state'];
                        $vendorAddressPostalCode = $items['activity_info']['activity_zip_code'];

                        $end_date = date("m/d/Y", strtotime($checkOutDate));
                        break;
                }
                if (!in_array($description, $arrayDescriptions)) {
                    $arrayDescriptions[] = $description;
                }
                $dashNumbre = array_keys($arrayDescriptions, $description)[0] + 1;
//                $dataItem['supplierName'] = $agency_info['agency_name'] . ":" .$clientLastName." ".$clientFirstName;
//                var_dump($clientLastName." ".$clientFirstName);
                $dataItem['supplierName'] = $itinerarxy_name;
//                var_dump($dataItem['supplierName']);
                $dataItem['startDate'] = $interval_start;//$checkInDate;

                $dataItem['endDate'] = $interval_end;//$end_date;
                $dataItem['retail'] = $rate_retail_before_tax;
                $dataItem['retailTax'] = $rate_tax_percent;
                $dataItem['retailFee'] = $rate_fee;
                $dataItem['net'] = $rate_net;
                $dataItem['netTax'] = $rate_tax_percent_net;
                $dataItem['netFee'] = $rate_fee_net;
                $dataItem['duration'] = $days;
                $dataItem['units'] = $units;
                $dataItem['stlRetail'] = "=(RC[-8]*RC[-2]*RC[-1])*RC[-7]+(RC[-6]*RC[-2]*RC[-1])";
                $dataItem['agencyCom'] = $rate_commission_percent;
                $dataItem['space'] = "=((RC[-10]*RC[-4]*RC[-3])*RC[-1])";
                $dataItem['soldPrice'] = "=(RC[-3]-RC[-1]";
                $dataItem['hotelCom'] = $rate_commission_percent_net;
                $dataItem['stlCost'] = "=((RC[-10]*RC[-7]*RC[-6])*RC[-1])";
                $dataItem['stlProfit'] = "=((RC[-11]*RC[-8]*RC[-7])*RC[-10])+(RC[-9]*RC[-8]*RC[-7])-((RC[-11]*RC[-8]*RC[-7])*RC[-2])";
                $dataItem['stlPercentage'] = "=(RC[-4]-RC[-1])";
                $dataItem['invoiceNumber'] = $invoiceId;
                $dataItem['type'] = $item_data['reservation_type'];
                $dataItem['invoiceNumber'] = $invoiceNumber;
                $dataItem['vendorName'] = $vendorName;
                $dataItem['description'] = $description;
                $dataItem['invoiceCreation'] = $invDate;
                $dataItem['agentName'] = $agentName;
                $dataItem['agentDept'] = $agentDept;
                $dataItem['agentOfc'] = $agentOffice;
                $dataItem['invoiceDash'] = $invoiceNumber . '-' . $dashNumbre;
                /*
                 *
                 *  $dataItem = array(
            'vendor' => 'vendor',//TBD Vendor
            'vendorAddressCity' => 'city',//TBD
            'vendorAddressCountry' => 'country',//TBD
            'vendorAddressLine1' => 'address',//TBD
            'vendorAddressState' => 'state',//TBD
            'vendorAddressPostalCode' => 'postal code',//TBD
            'dueDate' => 'due date',//TBD CHECKING DATE Y-m-d h:j:s
            'billNumber' => 'bull',//TBD invoice number plus dash
            'expenseLineAccount' => 'Cost of goods sold',//Concat Department and office
            'itemLineBillableStatus' => 'Billable',// Billable
            'itemLineCost' => 'item line cost',//TBD total price
            'itemLineCustomer' => 'item line customer',//TBD Agency name plus customer name
            'itemLineDescription' => 'item line description',//TBD Checkin and checkout dates
            'itemLineItem' => 'item line item',//TBD hotel, activity, trasnfers...
            'itemLineQuantity' => '1',//TBD just 1
            'transactionDate' => 'transaction date',//TBD Invoice date
        );
                 * */
                $expenseLineAccount = $agentOffice . '-' . $agentDept;
                $expenseLineAccountArray = array(
                    'Brazil-FIT' => 'Brazil - COGS',
                    'Canada-FIT' => 'Canada - COGS',
                    '-FIT' => 'FIT - COGS',
                    'Miami-FIT' => 'FIT - COGS',
                    'Miami-GROUP' => 'Groups - COGS',
                    'Miami-VIP' => 'VIP - COGS',
                );
                if (array_key_exists($expenseLineAccount, $expenseLineAccountArray)) {
                    $expenseLineAccount = $expenseLineAccountArray[$expenseLineAccount];
                }

                $reservationType = ucwords(str_replace('_', ' ', $items['reservation_type']));
                $reservationType .= ' ';
                $reservationType .= 'C';


                $reservationType .= ' - ';


                if ($agentOffice == 'Miami' || $agentOffice == '') {
                    if ($agentDept == "GROUP") {
                        $agentDept = "Groups";
                    }

                    $reservationType .= $agentDept;
                } else {
                    $reservationType .= $agentOffice;
                }

                if (array_key_exists($invoiceNumber . '-' . $dashNumbre, $dataArray)) {
                    $dataArray[$invoiceNumber . '-' . $dashNumbre]['itemLineCost'] += (($rate_net * $days * $units * $rate_tax_percent_net) + ($rate_fee_net * $days * $units));
                } else {
                    $dataItem['vendor'] = $vendorName;
                    $dataItem['vendorAddressCity'] = $vendorAddressCity;
                    $dataItem['vendorAddressCountry'] = $vendorAddressCountry;
                    $dataItem['vendorAddressLine1'] = $vendorAddressLine1;
                    $dataItem['vendorAddressState'] = $vendorAddressState;
                    $dataItem['vendorAddressPostalCode'] = $vendorAddressPostalCode;
                    $dataItem['dueDate'] = $checkInDate;
                    $dataItem['billNumber'] = str_replace('INT', 'I', str_replace('-', '', $invoiceNumber)) . '-' . $dashNumbre;
                    $dataItem['expenseLineAccount'] = $expenseLineAccount;
                    $dataItem['itemLineCost'] = ($rate_net * $days * $units * $rate_tax_percent_net) + ($rate_fee_net * $days * $units);
                    $dataItem['itemLineCustomer'] = $agency_info['agency_name'] . ' - ' . $clientFirstName . ' ' . $clientLastName;
                    $dataItem['itemLineDescription'] = $checkInDate . ' - ' . $checkOutDate;
                    $dataItem['itemLineItem'] = $reservationType;
                    $dataItem['transactionDate'] = $invDate;
                    $dataArray[$invoiceNumber . '-' . $dashNumbre] = $dataItem;
                }


                unset ($interval_start);

//                                var_dump($dataItem);
            }


        }

    }


//exit;

}

//var_dump($dataArray);

foreach ($dataArray as $dataItem) {

    $clean_code_postalCode = preg_replace('/[^\w]/', '', $dataItem['vendorAddressPostalCode']);

    $dataItem['vendorAddressPostalCode'] = $clean_code_postalCode;


    $clean_code_addressLine = preg_replace('/[^\w]/', '', $dataItem['vendorAddressLine1']);

    $dataItem['vendorAddressLine1'] = $clean_code_addressLine;

//    var_dump( $dataItem['vendorAddressPostalCode']);
    if ($dataItem['itemLineCost'] < 0) {
        continue;
    }
//     Output

//    var_dump( $dataItem['vendorAddressLine1']);

    $row++; // row counter

    $sheet->WriteString($row, 1, $dataItem['vendor'], $xml_styles['row_cell']);
    $sheet->WriteString($row, 3, $dataItem['vendorAddressCity'], $xml_styles['row_cell']);
    $sheet->WriteNumber($row, 4, $dataItem['vendorAddressCountry'], $xml_styles['row_cell']);

    $sheet->WriteNumber($row, 5, $dataItem['vendorAddressLine1'], $xml_styles['row_cell']);

    $sheet->WriteNumber($row, 11, $dataItem['vendorAddressState'], $xml_styles['row_cell']);
    $sheet->WriteNumber($row, 12, $dataItem['vendorAddressPostalCode'], $xml_styles['row_cell']);
    $sheet->WriteNumber($row, 13, $dataItem['dueDate'], $xml_styles['row_cell']);
    $sheet->WriteNumber($row, 14, $dataItem['billNumber'], $xml_styles['row_cell']);

    //    $sheet->WriteNumber($row, 21, $dataItem['expenseLineAccount'], $xml_styles['row_cell']);

    $sheet->WriteNumber($row, 28, $dataItem['itemLineBillableStatus'], $xml_styles['row_cell']);
    $sheet->WriteNumber($row, 30, $dataItem['itemLineCost'], $xml_styles['row_cell']);
    $sheet->WriteNumber($row, 31, $dataItem['itemLineCustomer'], $xml_styles['row_cell']);
    $sheet->WriteNumber($row, 32, $dataItem['itemLineDescription'], $xml_styles['row_cell']);
    $sheet->WriteNumber($row, 35, $dataItem['itemLineItem'], $xml_styles['row_cell']);
    $sheet->WriteNumber($row, 37, $dataItem['itemLineQuantity'], $xml_styles['row_cell']);
    $sheet->WriteNumber($row, 48, $dataItem['transactionDate'], $xml_styles['row_cell']);


    unset ($interval_start);

}

//exit;

$xml->sendHeaders();
$xml->writeData();

function get_xml_styles($xml)
{

    $output = array();

    $output['yellow_header'] = $xml->addStyle('yellow_header');
    $output['yellow_header']->bgColor('#fefe80');
    $output['yellow_header']->fontColor('Black');
    $output['yellow_header']->fontBold();

    $output['yellow_date'] = $xml->addStyle('yellow_date');
    $output['yellow_date']->numberFormatDate();
    $output['yellow_date']->bgColor('#FFF58C');
    $output['yellow_date']->fontColor('Black');
    $output['yellow_date']->alignHorizontal('Center');
    $output['yellow_date']->fontBold();

    // Headers
    $output['column_names'] = $xml->addStyle('green_header');
    $output['column_names']->bgColor('#EBF1DE');
    $output['column_names']->fontColor('Black');
    $output['column_names']->fontSize(8);
    $output['column_names']->border('All', 2);
    $output['column_names']->alignHorizontal('Center');
    $output['column_names']->fontBold();

    // Rows
    $output['accounting'] = $xml->addStyle('accounting');
    $output['accounting']->numberFormat('_(* #,##0.00_);_(* \(#,##0.00\);_(* &quot;-&quot;??_);_(@_)');
    $output['accounting']->border('All', 1);

    $output['percent'] = $xml->addStyle('percent');
    $output['percent']->numberFormat('0%');
    $output['percent']->border('All', 1);

    $output['percent_precise'] = $xml->addStyle('percent_precise');
    $output['percent_precise']->numberFormat('0.00%');
    $output['percent_precise']->border('All', 1);

    $output['date'] = $xml->addStyle('date');
    $output['date']->numberFormatDate();
    $output['date']->border('All', 1);

    $output['row_cell'] = $xml->addStyle('row_cell');
    $output['row_cell']->border('All', 1);


    // Subtotal
    $output['blue_footer'] = $xml->addStyle('blue_footer');
    $output['blue_footer']->bgColor('#c5d9ef');
    $output['blue_footer']->fontColor('Blue');
    $output['blue_footer']->fontBold();

    $output['blue_accounting'] = $xml->addStyle('blue_accounting');
    $output['blue_accounting']->numberFormat('_(* #,##0.00_);_(* \(#,##0.00\);_(* &quot;-&quot;??_);_(@_)');
    $output['blue_accounting']->bgColor('#c5d9ef');
    $output['blue_accounting']->fontColor('Blue');
    $output['blue_accounting']->fontBold();
    $output['blue_accounting']->border('All', 2);

    $output['red_footer'] = $xml->addStyle('red_footer');
    $output['red_footer']->bgColor('#F2DCDB');
    $output['red_footer']->fontColor('Red');
    $output['red_footer']->fontSize(8);
    $output['red_footer']->fontBold();

    $output['red_accounting'] = $xml->addStyle('red_accounting');
    $output['red_accounting']->numberFormat('_(* #,##0.00_);_(* \(#,##0.00\);_(* &quot;-&quot;??_);_(@_)');
    $output['red_accounting']->bgColor('#F2DCDB');
    $output['red_accounting']->fontColor('Red');
    $output['red_accounting']->fontSize(10);
    $output['red_accounting']->fontBold();
    $output['red_accounting']->border('All', 2);

    $output['green_footer'] = $xml->addStyle('green_footer');
    $output['green_footer']->bgColor('#EBF1DE');
    $output['green_footer']->fontColor('Green');
    $output['green_footer']->fontBold();

    $output['wire_fees'] = $xml->addStyle('wire_fees');
    $output['wire_fees']->bgColor('#EBF1DE');
    $output['wire_fees']->fontBold();
    $output['wire_fees']->border('All', 2);

    $output['invoice_total'] = $xml->addStyle('invoice_total');
    $output['invoice_total']->bgColor('#EBF1DE');
    $output['invoice_total']->fontBold();
    $output['invoice_total']->border('All', 2);

    $output['invoice_total_usd'] = $xml->addStyle('invoice_total_usd');
    $output['invoice_total_usd']->bgColor('#EBF1DE');
    $output['invoice_total_usd']->fontBold();
    $output['invoice_total_usd']->border('All', 2);
    $output['invoice_total_usd']->alignHorizontal('Center');


    $output['invoice_total_number'] = $xml->addStyle('invoice_total_number');
    $output['invoice_total_number']->numberFormat('_(* #,##0.00_);_(* \(#,##0.00\);_(* &quot;-&quot;??_);_(@_)');
    $output['invoice_total_number']->bgColor('#EBF1DE');
    $output['invoice_total_number']->fontBold();
    $output['invoice_total_number']->border('All', 2);

    $output['green_accounting'] = $xml->addStyle('green_accounting');
    $output['green_accounting']->numberFormat('_(* #,##0.00_);_(* \(#,##0.00\);_(* &quot;-&quot;??_);_(@_)');
    $output['green_accounting']->bgColor('#EBF1DE');
    $output['green_accounting']->fontColor('Green');
    $output['green_accounting']->fontBold();
    $output['green_accounting']->border('All', 2);

    $output['green_percent'] = $xml->addStyle('green_percent');
    $output['green_percent']->numberFormat('0.00%');
    $output['green_percent']->bgColor('#EBF1DE');
    $output['green_percent']->fontColor('Green');
    $output['green_percent']->fontBold();

    $output['subtotal'] = $xml->addStyle('subtotal');
    $output['subtotal']->bgColor('#C5D9EF');
    $output['subtotal']->fontBold();
    $output['subtotal']->border('All', 2);

    $output['stl_percent_total'] = $xml->addStyle('stl_percent_total');
    $output['stl_percent_total']->numberFormat('0.00%');
    $output['stl_percent_total']->bgColor('#EBF1DE');
    $output['stl_percent_total']->fontBold();
    $output['stl_percent_total']->border('All', 2);


    return $output;
}

/**
 * @param $search_string
 * @param $offset
 * @param $limit
 * @param $sort_by
 * @param $visible_paid_status
 * @return array
 */


function GetInvoices($search_string, $offset, $limit, $sort_by, $visible_paid_status)
{


    global $dashboard_pdo;

    $output = array();

    $search = ParseSearchString($search_string);

    $query = "  SELECT
                inv.invoice_id,
                inv.invoice_number,
                inv.quote_id,
                inv.pay_status,
                inv.invoice_status,
                inv.amount_invoiced,
                inv.amount_paid,
                inv.revision,
                inv.payment_due,
                inv.payment_type,
                inv.payment_fee,
                inv.payment_fee_type,
                inv.creation_date,
                q.quote_title,
                q.agency_id,
                q.created_by,
                q.quote_data,
                q.agency_name,
                u.user_first_name,
                inv.commission_agency

                FROM oti_invoices AS inv
                INNER JOIN quotes AS q ON q.quote_id = inv.quote_id
                INNER JOIN oti_users AS u ON q.created_by = u.username
                INNER JOIN invoice_items AS items ON items.invoice_id = inv.invoice_id

                WHERE ( {$search[ 'field_name' ]} {$search[ 'search_type' ]} :search_string
                OR invoice_number LIKE :search_string ) ";

    switch ($visible_paid_status) {
        case 'hide':
            $query .= "AND inv.pay_status != 'PAID' ";
            break;
    }


    $query .= "AND inv.creation_date > '2021-05-01'";
    $query .= "AND inv.creation_date < '2021-06-01'";


    $query .= "GROUP BY inv.invoice_id ";


    switch ($sort_by) {

        case 'payment_due_date':
            $query .= "ORDER BY inv.payment_due ASC";
            break;
        case 'check_in_date':
            $query .= "ORDER BY items.check_in_date ASC ";
            break;
        case 'check_out_date':
            $query .= "ORDER BY items.check_out_date ASC ";
            break;
        case 'revision':
            $query .= "ORDER BY inv.revision DESC ";
            break;
        case 'pending_approval':
            $query .= "ORDER BY inv.invoice_status ASC, inv.invoice_NUMBER DESC ";
            break;
        default:
            $query .= "ORDER BY inv.invoice_number DESC";
            break;
    }

    if ($limit) {

        $query .= " LIMIT :offset, :limit ";

    }


    $result = $dashboard_pdo->prepare($query);


    // Handle the search type loose or strict
    if ($search['search_type'] == 'LIKE') {

        $result->bindValue(':search_string', "%{$search[ 'field_value' ]}%", PDO::PARAM_STR);
    } else {
        $result->bindValue(':search_string', $search['field_value'], PDO::PARAM_STR);
    }

    if ($limit) {
        $result->bindValue(':offset', $offset, PDO::PARAM_INT);
        $result->bindValue(':limit', $limit, PDO::PARAM_INT);
    }
    $result->execute();


    $result_array = array();

    while ($row = $result->fetch(PDO::FETCH_ASSOC)) {

        $output[$row['invoice_id']]['quote_id'] = $row['quote_id'];
        $output[$row['invoice_id']]['agency_id'] = $row['agency_id'];
        $output[$row['invoice_id']]['quote_title'] = $row['quote_title'];
        $output[$row['invoice_id']]['quote_data'] = unserialize($row['quote_data']);
        $output[$row['invoice_id']]['invoice_number'] = $row['invoice_number'];
        $output[$row['invoice_id']]['created_by'] = $row['user_first_name'];
        $output[$row['invoice_id']]['pay_status'] = $row['pay_status'];
        $output[$row['invoice_id']]['invoice_status'] = $row['invoice_status'];
        $output[$row['invoice_id']]['pay_status_word'] = ($row['pay_status'] == 'PPAY') ? 'Pending' : 'Paid';
        $output[$row['invoice_id']]['amount_invoiced'] = $row['amount_invoiced'];
        $output[$row['invoice_id']]['creation_date'] = $row['creation_date'];

        $total_amount_invoiced = $row['amount_invoiced'] + CalculatePaymentFees($row['amount_invoiced'], $row['payment_fee_type'], $row['payment_fee']);

        $output[$row['invoice_id']]['total_amount_invoiced'] = $total_amount_invoiced;
        $output[$row['invoice_id']]['amount_paid'] = $row['amount_paid'];
        $output[$row['invoice_id']]['amount_balance'] = $total_amount_invoiced - $row['amount_paid'];
        $output[$row['invoice_id']]['revision'] = $row['revision'];
        $output[$row['invoice_id']]['payment_due'] = $row['payment_due'];
        $output[$row['invoice_id']]['checkbox_status'] = ($row['pay_status'] == 'PAID') ? 'disabled' : '';
        $output[$row['invoice_id']]['payment_type'] = $row['payment_type'];
        $output[$row['invoice_id']]['payment_fee'] = $row['payment_fee'];
        $output[$row['invoice_id']]['payment_fee_type'] = $row['payment_fee_type'];
        $output[$row['invoice_id']]['agency_name'] = $row['agency_name'];

        $query = "      SELECT
                        items.*,
                        h.hotel_name,
                        os.supplier_name
                        FROM oti_invoices AS i
                        INNER JOIN quotes AS q ON q.quote_id = i.quote_id
                        INNER JOIN invoice_items AS items ON items.invoice_id = i.invoice_id
                        LEFT JOIN " . DB_MAIN . ".oti_hotels AS h ON h.hotel_id = items.hotel_id
                        LEFT JOIN oti_suppliers AS os ON items.supplier_id = os.supplier_id

                        WHERE i.invoice_id = :invoice_id ";

        $stmt = $dashboard_pdo->prepare($query);
        $stmt->bindValue(':invoice_id', $row['invoice_id']);

        $stmt->execute();


        while ($item_row = $stmt->fetch(PDO::FETCH_ASSOC)) {
            $result_array[] = $item_row;
        }
    }


    $old_quote_id = false;

    foreach ($result_array as $row) {


        $reservation_data = unserialize($row['reservation']);

        if ($old_quote_id != $row['invoice_id']) {
            // Next quote
            $total_retail = 0;
            $total_commission = 0;
        }

        // Load a default color if not set
        if (!isset ($reservation_data['item_color'])) {
            $reservation_data['item_color'] = '#ccc';
        }

        $old_quote_id = $row['invoice_id'];

        $output[$row['invoice_id']]['items'][$row['invoice_item_id']]['reservation'] = $reservation_data;
        $output[$row['invoice_id']]['items'][$row['invoice_item_id']]['reservation_type'] = $row['reservation_type'];
        $output[$row['invoice_id']]['items'][$row['invoice_item_id']]['created_by'] = $row['created_by'];
        $output[$row['invoice_id']]['items'][$row['invoice_item_id']]['cancellation_due'] = $row['cancellation_due'];
        $output[$row['invoice_id']]['items'][$row['invoice_item_id']]['modification_type'] = $row['modification_type'];

        if (($row['reservation_type'] == 'hotel') or ($row['reservation_type'] == 'hotel_service') or ($row['reservation_type'] == 'resort_fee')) {
            $output[$row['invoice_id']]['items'][$row['invoice_item_id']]['reservation']['hotel_info']['hotel_name'] = $row['hotel_name'];
        } else {
            $output[$row['invoice_id']]['items'][$row['invoice_item_id']]['reservation']['supplier_name'] = $row['supplier_name'];
        }

        try {
            $total_retail += ($reservation_data['rates_data']['total']['rate_retail_after_tax'] * $reservation_data['request']['units']);
            $total_commission += ($reservation_data['rates_data']['total']['rate_commission'] * $reservation_data['request']['units']);
        } catch (Exception $e) {
//            print_r($reservation_data);
        }
        $output[$row['invoice_id']]['total_retail'] = round($total_retail, 2);
        $output[$row['invoice_id']]['total_commission'] = round($total_commission, 2);
    }


    // Sort the items in this quote by their position
    foreach ($output as $invoice_id => $invoice) {
        uasort($output[$invoice_id]['items'], 'sort_reservation_by_position');
    }


    return $output;
}

/**
 * @param $status_code
 * @return string
 */
function GetConfirmationStatusWord($status_code)
{

    switch ($status_code) {
        case 'CF':
            $status_word = 'Confirmed';
            break;
        case 'CP':
            $status_word = 'Pending Conf.';
            break;
        case 'CC':
            $status_word = 'Cancelled';
            break;
        case 'CH':
            $status_word = 'On Hold';
            break;
    }

    return $status_word;
}

/**
 * @param $status_code
 * @return string
 */
function GetPaymentStatusWord($status_code)
{

    switch ($status_code) {
        case 'PPAY':
            $status_word = 'Pending Payment';
            break;
        case 'PAID':
            $status_word = 'Paid';
            break;
        case 'PREF':
            $status_word = 'Pending Refund';
            break;
        case 'REFD':
            $status_word = 'Refunded';
            break;
        default:
            $status_word = 'Pending Payment';
            break;
    }

    return $status_word;
}

/**
 * @param $search_string
 * @return array
 */
function ParseSearchString($search_string)
{

    $output = array();

    $output['search_type'] = 'LIKE';    // set the default search type
    // Manage search
    if (strstr($search_string, ':')) {
        $exploded_search = explode(':', $search_string);

        switch (strtolower($exploded_search[0])) {
            case 'quote_id':
                $output['field_name'] = 'q.quote_id';
                $output['search_type'] = '=';   // Set strict search type
                break;
            case 'quote_title':
                $output['field_name'] = 'q.quote_title';
                break;
            case 'agency_id':
                $output['field_name'] = 'q.agency_id';
                $output['search_type'] = '=';   // Set strict search type
                break;
            case 'invoice_id':
                $output['field_name'] = 'i.invoice_id';
                $output['search_type'] = '=';
                break;
            default:
                $output['field_name'] = 'q.quote_title';
                break;
        }

        $output['field_value'] = $exploded_search[1];
    } else {
        $output['field_name'] = 'q.quote_title';
        $output['field_value'] = $search_string;
    }

    // trim spaces
    $output['field_value'] = trim($output['field_value']);


    return $output;
}

function GetDuration($startingDate, $endingDate)
{
    $startingMonth = $startingDate[5] . $startingDate[6];
    $startingDay = $startingDate[8] . $startingDate[9];
    $startingYear = $startingDate[0] . $startingDate[1] . $startingDate[2] . $startingDate[3];

    $endingMonth = $endingDate[5] . $endingDate[6];
    $endingDay = $endingDate[8] . $endingDate[9];
    $endingYear = $endingDate[0] . $endingDate[1] . $endingDate[2] . $endingDate[3];

    if ($startingMonth == $endingMonth) {

        $duration = $endingDay - $startingDay;
    }

    if ($startingMonth != $endingMonth) {
        $remainder = 31 - $startingDay;

        $duration = $endingDay + $remainder;
    }

    if ($duration == 0) {
        $duration = 1;
    }

    return $duration;

}