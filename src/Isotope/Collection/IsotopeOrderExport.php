<?php

/**
 * IsotopeOrderExport
 *
 * @copyright  inszenium 2016 <https://inszenium.de>
 * @author     Kirsten Roschanski <kirsten.roschanski@inszenium.de>
 * @package    IsotopeOrderExport
 * @license    LGPL 
 * @link       https://github.com/inszenium/isotope-export
 * @filesource
 */

namespace Isotope\Collection;
use Isotope\Isotope;
use Isotope\Interfaces\IsotopeProduct;

/**
 * Class IsotopeOrderExport
 * Provide miscellaneous methods that are used by the data configuration array.
 */
class IsotopeOrderExport extends \Backend
{

  /**
   * An array with contains the fields
   * @var array
   */
  protected $arrContent = array();


  /**
   * An array with the header fields
   * @var array
   */
  protected $arrHeaderFields = array();


  /**
   * The field delimiter
   * @var string
   */
  protected $strDelimiter = '"';


  /**
   * The field seperator
   * @var string
   */
  protected $strSeperator = ';';


  /**
   * The line end
   * @var string
   */
  protected $strLineEnd = "\r\n";


  /**
   * Import an Isotope object
   */
  public function __construct()
  {
    parent::__construct();
    \System::loadLanguageFile('countries');
  }


  /**
   * Gibt ein Mapping von product_id => tax_class zurück (nur erster Eintrag pro Produkt).
   *
   * @return array
  */
  protected function getTaxClassMapping(): array
  {
      $arrTaxClasses = [];
  
      $objTaxData = \Database::getInstance()->query("SELECT pid, tax_class FROM tl_iso_product_price");
  
      while ($objTaxData->next()) {
          // Nur speichern, wenn pid noch nicht vorhanden ist
          if (!isset($arrTaxClasses[$objTaxData->pid])) {
              $arrTaxClasses[$objTaxData->pid] = $objTaxData->tax_class;
          }
      }
  
      return $arrTaxClasses;
  }


  /**
   * Generate the csv file and send it to the browser
   *
   * @param void
   * @return void
   */
  public function saveToBrowser()
  {
    if (count($this->arrContent) < 1) {
      $strRequest = ampersand(str_replace('&key=export_order', '', $this->Environment->request));
      $strRequest = ampersand(str_replace('&excel=true', '', $strRequest));

      return '<div id="tl_buttons">
          <a href="' . $strRequest . '" class="header_back" title="' . specialchars($GLOBALS['TL_LANG']['MSC']['backBT']) . '">' . $GLOBALS['TL_LANG']['MSC']['backBT'] . '</a>
          </div>
          <p class="tl_gerror">' . $GLOBALS['TL_LANG']['MSC']['noOrders'] . '</p>';
    }


    $strContent = $this->prepareContent();

    // Ensure UTF-8 encoding and add BOM for Excel compatibility
    $bom = chr(0xEF) . chr(0xBB) . chr(0xBF);
    $strContent = $bom . $strContent; // Prepend BOM to the content to ensure UTF-8

    header('Content-Transfer-Encoding: binary');
    header('Content-Disposition: attachment; filename="isotope_items_export_' . $this->parseDate($GLOBALS['TL_CONFIG']['dateFormat'], time()) . '_' . $this->parseDate($GLOBALS['TL_CONFIG']['timeFormat'], time()) . '.csv"');
    header('Content-Length: ' . strlen($strContent));
    header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
    header('Pragma: public');
    header('Content-Type: text/html; charset=utf-8');

    header('Expires: 0');

    echo $strContent;
    exit;
  }


  /**
   * Export orders and send them to browser as file
   * @param DataContainer
   * @return string
   */
  public function exportOrders()
  {
    if ($this->Input->get('key') != 'export_orders') {
      return '';
    }

    $csvHead = &$GLOBALS['TL_LANG']['tl_iso_product_collection']['csv_head'];
    $arrKeys = ['order_id', 'date', 'company', 'lastname', 'firstname', 'street', 'postal', 'city', 'country', 'phone', 'email', 'items', 'subTotal', 'grandTotal'];

    // Fetch the current year (e.g., '25' for 2025)
    $currentYear = date('y');  // 'y' gives two digits of the current year (e.g., '25' for 2025)

    // Fetch all order items
    $objOrderItems = \Database::getInstance()->query("SELECT * FROM tl_iso_product_collection_item");

    $arrOrderItems = [];
    $arrOrderSKUs = [];
    $arrOrderPrices = [];

    while ($objOrderItems->next()) {
      $orderId = $objOrderItems->pid;
      if (!isset($arrOrderItems[$orderId])) {
        $arrOrderItems[$orderId] = [];
        $arrOrderSKUs[$orderId] = [];
        $arrOrderPrices[$orderId] = [];
      }

      $arrOrderItems[$orderId][] = html_entity_decode(
        $objOrderItems->quantity . " x " . strip_tags($objOrderItems->name) . " [" . $objOrderItems->sku . "] " .
        " á " . strip_tags(Isotope::formatPrice($objOrderItems->price)) .
        " (" . strip_tags(Isotope::formatPrice($objOrderItems->quantity * $objOrderItems->price)) . ")"
      );

      $arrOrderSKUs[$orderId][] = $objOrderItems->sku;
      $arrOrderPrices[$orderId][] = Isotope::formatPrice($objOrderItems->price * $objOrderItems->quantity);
    }

    // Determine max number of items in any order
    $maxItems = max(array_map('count', $arrOrderSKUs));

    // Add SKU and price columns dynamically
    for ($i = 1; $i <= $maxItems; $i++) {
      $arrKeys[] = "Artikelnummer $i";
      $arrKeys[] = "Preis Gesamt $i";
    }

    foreach ($arrKeys as $v) {
      $this->arrHeaderFields[$v] = $csvHead[$v] ?? $v;
    }

    // Fetch orders
    $objOrders = \Database::getInstance()->query("SELECT *, tl_iso_product_collection.id as collection_id FROM tl_iso_product_collection, tl_iso_address WHERE tl_iso_product_collection.billing_address_id = tl_iso_address.id AND (document_number != '' OR document_number IS NOT NULL) AND document_number LIKE '$currentYear%'  -- Filter by current year in document_number ORDER BY document_number ASC");

    if (null === $objOrders) {
      return '<div id="tl_buttons">
            <a href="' . ampersand(str_replace('&key=export_order', '', $this->Environment->request)) . '" class="header_back" title="' . specialchars($GLOBALS['TL_LANG']['MSC']['backBT']) . '">' . $GLOBALS['TL_LANG']['MSC']['backBT'] . '</a>
            </div>
            <p class="tl_gerror">' . $GLOBALS['TL_LANG']['MSC']['noOrders'] . '</p>';
    }

    while ($objOrders->next()) {
      if (!isset($arrOrderItems[$objOrders->collection_id])) {
        continue;
      }

      // Prepare SKU and price columns
      $skuColumns = array_pad($arrOrderSKUs[$objOrders->collection_id], $maxItems, ''); // Fill missing columns with empty strings
      $priceColumns = array_pad($arrOrderPrices[$objOrders->collection_id], $maxItems, '');

      $this->arrContent[] = array_merge([
        //'status' => $objOrders->order_status,
        'order_id' => $objOrders->document_number,
        'date' => $this->parseDate($GLOBALS['TL_CONFIG']['datimFormat'], $objOrders->locked),
        'company' => $objOrders->company,
        'lastname' => $objOrders->lastname,
        'firstname' => $objOrders->firstname,
        'street' => $objOrders->street_1,
        'postal' => $objOrders->postal,
        'city' => $objOrders->city,
        'country' => $GLOBALS['TL_LANG']['CNT'][$objOrders->country],
        'phone' => $objOrders->phone,
        'email' => $objOrders->email,
        'items' => implode(' ', $arrOrderItems[$objOrders->collection_id]),

        //'subTotal' => number_format($objOrders->subTotal, 2, ',', ''),
        'subTotal' => number_format($objOrders->tax_free_subtotal, 2, ',', ''),
        'grandTotal' => number_format($objOrders->total, 2, ',', ''),
      ], array_merge(...array_map(null, $skuColumns, $priceColumns))); // Merge SKU and price columns alternatively
    }

    // Output
    $this->saveToBrowser();
  }



  /**
   * Export orders and send them to browser as file
   * @param DataContainer
   */
  public function exportItems()
  {
    if ($this->Input->get('key') != 'export_items') {
      return '';
    }

    $csvHead = &$GLOBALS['TL_LANG']['tl_iso_product_collection']['csv_head'];
    $arrKeys = array('order_id', 'date', 'company', 'lastname', 'firstname', 'street', 'postal', 'city', 'country', 'phone', 'email', 'count', 'item_sku', 'item_name', 'item_price', 'item_price_with_tax', 'tax_rate', 'tax', 'final_price', 'sum', 'tax_class');

    foreach ($arrKeys as $v) {
      $this->arrHeaderFields[$v] = $csvHead[$v];
    }

    // Fetch orders
    $objOrders = \Database::getInstance()->query("SELECT *, tl_iso_product_collection.id as collection_id FROM tl_iso_product_collection, tl_iso_address WHERE tl_iso_product_collection.billing_address_id = tl_iso_address.id AND (document_number != '' OR document_number IS NOT NULL) ORDER BY document_number ASC");

    if (null === $objOrders) {
      return '<div id="tl_buttons">
                    <a href="' . ampersand(str_replace('&key=export_order', '', $this->Environment->request)) . '" class="header_back" title="' . specialchars($GLOBALS['TL_LANG']['MSC']['backBT']) . '">' . $GLOBALS['TL_LANG']['MSC']['backBT'] . '</a>
                    </div>
                    <p class="tl_gerror">' . $GLOBALS['TL_LANG']['MSC']['noOrders'] . '</p>';
    }

    // Fetch order items
    $objOrderItems = \Database::getInstance()->query("SELECT * FROM tl_iso_product_collection_item");

    $arrOrderItems = array();
    while ($objOrderItems->next()) {
      $arrConfig = deserialize($objOrderItems->configuration);
      $strConfig = '';
      if (is_array($arrConfig)) {
        foreach ($arrConfig as $key => $value) {
          if (strlen($strConfig) > 1) {
            $strConfig .= PHP_EOL;
          }
          $arrValues = deserialize($value);
          $strConfig .= \Isotope\Translation::get($key) . ": " . (is_array($arrValues) ? implode(",", $arrValues) : \Isotope\Translation::get($value));
        }
      }

      // Store order items by order ID
      $arrOrderItems[$objOrderItems->pid][] = array(
        'count' => $objOrderItems->quantity,
        'item_sku' => html_entity_decode($objOrderItems->sku),
        'item_name' => strip_tags(html_entity_decode($objOrderItems->name)),
        'item_price' => Isotope::formatPrice($objOrderItems->price),
        'configuration' => strip_tags(html_entity_decode($strConfig)),
        'sum' => Isotope::formatPrice($objOrderItems->quantity * $objOrderItems->price),
        'product_id' => $objOrderItems->product_id,
      );
    }

    // Fetch surcharges (shipping and tax)
    $objSurcharges = \Database::getInstance()->query("SELECT * FROM tl_iso_product_collection_surcharge WHERE type IN ('shipping', 'tax')");

    $arrSurcharges = [];
    while ($objSurcharges->next()) {
      $arrSurcharges[$objSurcharges->pid][$objSurcharges->type] = [
        'price' => $objSurcharges->price,
        'total_price' => $objSurcharges->total_price,
        'tax' => $objSurcharges->tax
      ];
    }

// Include shipping surcharge and apply correct tax based on tax_class
foreach ($arrSurcharges as $pid => $surcharge) {
  if (isset($surcharge['shipping'])) {
    $shipping_total_price = (float) str_replace(',', '.', $surcharge['shipping']['total_price']);  // Gross
    $shipping_tax = isset($surcharge['shipping']['tax']) ? (float) str_replace(',', '.', $surcharge['shipping']['tax']) : 0;

    // Determine tax rate based on tax_class
    $tax_rate = 0.00; // default fallback
    if (isset($surcharge['shipping']['tax_class'])) {
      switch ((int) $surcharge['shipping']['tax_class']) {
        case 2:
          $tax_rate = 0.19;
          break;
        case 4:
          $tax_rate = 0.07;
          break;
        default:
          $tax_rate = 0.00;
      }
    }

    // Calculate price including tax
    $final_price = $shipping_total_price;
    if ($tax_rate > 0) {
      $final_price += $shipping_total_price * $tax_rate;
    }

    $final_price = round($final_price, 2);

    // Add the shipping surcharge as an item with tax applied
    $arrOrderItems[$pid][] = array(
      'count' => 1,
      'item_sku' => '',
      'item_name' => 'Versandkosten',
      'item_price' => Isotope::formatPrice($shipping_total_price),
      'tax_rate' => $tax_rate * 100,  // Show as 0, 7, or 19
      'item_price_with_tax' => Isotope::formatPrice($final_price),
      'sum' => 5,
    );
  }
}

    $taxClassMap = $this->getTaxClassMapping();

    // Compile data for export
    while ($objOrders->next()) {
       
      if (!isset($arrSurcharges[$objOrders->collection_id]['shipping'])) {
        continue;
      }
      // Check if the order_id (Bestell-Id) is not empty and it has shipping surcharge items
      if (!isset($arrOrderItems[$objOrders->collection_id]) || empty($objOrders->document_number)) {
        continue;  // Skip this order if order_id is empty or no shipping surcharge exists
      }
      foreach ($arrOrderItems[$objOrders->collection_id] as $item) {
        // tax_rate auf Basis von tax_class berechnen
        $tax_rate = 0;
        $tax_class = $item['tax_class'] ?? ($item['product_id'] ? ($taxClassMap[$item['product_id']] ?? '') : '');
        
        switch ((int) $tax_class) {
          case 2:
            $tax_rate = 0.19;
            break;
          case 4:
            $tax_rate = 0.07;
            break;
          default:
            $tax_rate = 0.00;
        }

      // Calculate Item Tax and Item Price with Tax
      $item_price = (float) strtr($item['item_price'], array('.' => '', ',' => '.'));
      $item_tax = (float) $item_price * $tax_rate;

      $item_price_with_tax = (float) $item_price + $item_tax;

      $formatted_item_price_with_tax = number_format($item_price_with_tax, 2, ',', '.');
      $formatted_item_tax = number_format($item_tax, 2, ',', '.');

      $final_price = $item_price_with_tax * $item['count'];
      $final_price = number_format($final_price, 2, ',', '.');

      $sum = $item['sum'];
      $price = $item['item_price'];


      //Add SKU to Shipping
      $sku = $item['item_sku'];

      if($item['item_name'] == "Versandkosten") {
        $sku = "84160";
      }

      //Add Minus to Stornorechnungen
      if($objOrders->order_status == "5") {
        $formatted_item_price_with_tax = "-" . $formatted_item_price_with_tax;
        $formatted_item_tax = "-" . $formatted_item_tax;
        $final_price = "-" . $final_price;
        $sum = "-" . $sum;
        $price = "-" . $price;
      }

        // Add product item to export content
        $this->arrContent[] = array(
          'order_id' => $objOrders->document_number,
          'date' => $this->parseDate($GLOBALS['TL_CONFIG']['datimFormat'], $objOrders->locked),
          'company' => $objOrders->company,
          'lastname' => $objOrders->lastname,
          'firstname' => $objOrders->firstname,
          'street' => $objOrders->street_1,
          'postal' => $objOrders->postal,
          'city' => $objOrders->city,
          'country' => $GLOBALS['TL_LANG']['CNT'][$objOrders->country],
          'phone' => $objOrders->phone,
          'email' => $objOrders->email,
          'count' => $item['count'],
          'item_sku' => $sku,
          'item_name' => $item['item_name'],
          'item_price' => $price,
          'item_price_with_tax' => $formatted_item_price_with_tax, // New column for price with tax
          'tax_rate' => $tax_rate * 100, // Tax rate as 0, 7, or 19
          'tax' => $formatted_item_tax ?? '',
          'final_price' => $final_price ?? '',
          'sum' => $sum,
          'tax_class' => isset($item['product_id']) ? ($taxClassMap[$item['product_id']] ?? '') : '',
        );

      }
    }

    // Output CSV file
    $this->saveToBrowser();
  }




  /**
   * Prepare the given array and build the content stream
   *
   * @param void
   * @return string
   */
  public function prepareContent()
  {
    $strCsv = '';
    $arrData = array();

    // add the header fields if there are some
    if (count($this->arrHeaderFields) > 0) {
      $arrData = array($this->arrHeaderFields);
    }

    // add all other elements
    foreach ($this->arrContent as $k => $v) {
      //TODO: maybe find a better solution
      $arrData[] = $v;
    }


    // build the csv string
    foreach ((array) $arrData as $arrRow) {
      array_walk($arrRow, array($this, 'escapeRow'));
      $strCsv .= $this->strDelimiter . implode($this->strDelimiter . $this->strSeperator . $this->strDelimiter, $arrRow) . $this->strDelimiter . $this->strLineEnd;
    }

    // add the excel support if requested
    if ($this->Input->get('excel')) {
      $strCsv = chr(255) . chr(254) . mb_convert_encoding($strCsv, 'UTF-16LE', 'UTF-8');
    }

    return $strCsv;
  }

  /**
   * Escape a row
   *
   * @param mixed &$varValue
   * @return void
   */
  public function escapeRow(&$varValue)
  {
    $varValue = str_replace('"', '""', $varValue);
  }
}
