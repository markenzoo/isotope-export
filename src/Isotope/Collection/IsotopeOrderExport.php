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
   * Generate the csv file and send it to the browser
   *
   * @param void
   * @return void
   */
  public function saveToBrowser()
  {
	if ( count($this->arrContent) < 1 ) {
	  $strRequest = ampersand(str_replace('&key=export_order', '', $this->Environment->request));	
	  $strRequest = ampersand(str_replace('&excel=true', '', $strRequest));	
	  
      return '<div id="tl_buttons">
          <a href="'.$strRequest.'" class="header_back" title="'.specialchars($GLOBALS['TL_LANG']['MSC']['backBT']).'">'.$GLOBALS['TL_LANG']['MSC']['backBT'].'</a>
          </div>
          <p class="tl_gerror">'. $GLOBALS['TL_LANG']['MSC']['noOrders'] .'</p>';
    }     
	  
    
    $strContent = $this->prepareContent();

    // Ensure UTF-8 encoding and add BOM for Excel compatibility
    $bom = chr(0xEF) . chr(0xBB) . chr(0xBF);
    $strContent = $bom . $strContent; // Prepend BOM to the content to ensure UTF-8

    header('Content-Transfer-Encoding: binary');
    header('Content-Disposition: attachment; filename="isotope_items_export_' . $this->parseDate($GLOBALS['TL_CONFIG']['dateFormat'], time()) . '_' . $this->parseDate($GLOBALS['TL_CONFIG']['timeFormat'], time()) .'.csv"');
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
    $arrKeys = ['status', 'order_id', 'date', 'company', 'lastname', 'firstname', 'street', 'postal', 'city', 'country', 'phone', 'email', 'items', 'subTotal', 'grandTotal'];

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
      $arrKeys[] = "Artikelnummer_$i";
      $arrKeys[] = "Preis_Gesamt_$i";
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
        'status' => $objOrders->order_status,
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

<<<<<<< HEAD
        //'subTotal' => number_format($objOrders->subTotal, 2, ',', ''),
        'subTotal' => number_format($objOrders->tax_free_subtotal, 2, ',', ''),
        'grandTotal' => number_format($objOrders->total, 2, ',', ''),
      ], array_merge(...array_map(null, $skuColumns, $priceColumns))); // Merge SKU and price columns alternatively
=======
     // Format as number without prepending quote
    $subTotalFormatted = number_format($subTotal, 2, ',', '');  // European format (comma for decimal)
    $taxTotalFormatted = number_format($taxTotal, 2, ',', '');    // European format
    $grandTotalFormatted = number_format($grandTotal, 2, ',', ''); // European format
	    
      $this->arrContent[] = array(   
  'status'             => $objOrders->order_status, 
        'order_id'      => $objOrders->document_number,
        'date'          => $this->parseDate($GLOBALS['TL_CONFIG']['datimFormat'], $objOrders->locked),
        'company'       => $objOrders->company, 
        'lastname'      => $objOrders->lastname, 
        'firstname'     => $objOrders->firstname,
        'street'        => $objOrders->street_1, 
        'postal'        => $objOrders->postal, 
        'city'          => $objOrders->city, 
        'country'       => $GLOBALS['TL_LANG']['CNT'][$objOrders->country],
        'phone'         => $objOrders->phone, 
        'email'         => $objOrders->email,
        'items'         => $arrOrderItems[$objOrders->collection_id],
        'subTotal'       => $subTotalFormatted,  
        'taxTotal'       => $taxTotalFormatted,  
        'grandTotal'     => $grandTotalFormatted, 
        'item_sku'       => $arrOrderSKUs[$objOrders->collection_id],

      );

      //$this->arrContent += ['item_sku' => $arrOrderSKUs[$objOrders->collection_id]];


>>>>>>> e09c27852cc62dcbf9b73b7c00a93df766b768a6
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
    $arrKeys = array('status', 'order_id', 'date', 'company', 'lastname', 'firstname', 'street', 'postal', 'city', 'country', 'phone', 'email', 'count', 'item_sku', 'item_name', 'item_configuration', 'item_price', 'sum');
   
    foreach ($arrKeys as $v) {
      $this->arrHeaderFields[$v] = $csvHead[$v];
    } 
   
    $objOrders = \Database::getInstance()->query("SELECT *, tl_iso_product_collection.id as collection_id FROM tl_iso_product_collection, tl_iso_address WHERE tl_iso_product_collection.billing_address_id = tl_iso_address.id AND ( document_number != '' OR document_number IS NOT NULL) ORDER BY document_number ASC");

    if (null === $objOrders) {
      return '<div id="tl_buttons">
          <a href="'.ampersand(str_replace('&key=export_order', '', $this->Environment->request)).'" class="header_back" title="'.specialchars($GLOBALS['TL_LANG']['MSC']['backBT']).'">'.$GLOBALS['TL_LANG']['MSC']['backBT'].'</a>
          </div>
          <p class="tl_gerror">'. $GLOBALS['TL_LANG']['MSC']['noOrders'] .'</p>';
    }
    
    $objOrderItems = \Database::getInstance()->query("SELECT * FROM tl_iso_product_collection_item");      
    
    $arrOrderItems = array();
    while ($objOrderItems->next()) {  
	  $arrConfig = deserialize($objOrderItems->configuration);
	  $strConfig = '';
	  if(is_array($arrConfig)) {
		foreach ($arrConfig as $key => $value) {
		  if( strlen($strConfig) > 1 ) {
			$strConfig .= PHP_EOL;
		  }	
		  $arrValues = deserialize($value);
		  $strConfig .= \Isotope\Translation::get($key) . ": " . (is_array($arrValues) ? implode(",", $arrValues) : \Isotope\Translation::get($value));
		}	
	  }	
		
		                
      $arrOrderItems[$objOrderItems->pid][] = array
      (
        'count'         => $objOrderItems->quantity,
        'item_sku'      => html_entity_decode( $objOrderItems->sku ),
        'item_name'     => strip_tags(html_entity_decode($objOrderItems->name)),
        'item_price'    => strip_tags(html_entity_decode(Isotope::formatPrice($objOrderItems->price))),
        'configuration' => strip_tags(html_entity_decode($strConfig)),
        'sum'           => strip_tags(html_entity_decode(Isotope::formatPrice($objOrderItems->quantity * $objOrderItems->price))),    
      );    
    }

    while ($objOrders->next()) {
      if( isset($arrOrderItems) && is_array($arrOrderItems) && !array_key_exists($objOrders->collection_id, $arrOrderItems) ) { continue; }
  
      foreach ($arrOrderItems[$objOrders->collection_id] as $item) {
        $this->arrContent[] = array(
	  
          'order_id'           => $objOrders->document_number,
          'date'               => $this->parseDate($GLOBALS['TL_CONFIG']['datimFormat'], $objOrders->locked),
          'company'            => $objOrders->company, 
          'lastname'           => $objOrders->lastname, 
          'firstname'          => $objOrders->firstname,
          'street'             => $objOrders->street_1, 
          'postal'             => $objOrders->postal, 
          'city'               => $objOrders->city, 
          'country'            => $GLOBALS['TL_LANG']['CNT'][$objOrders->country],
          'phone'              => $objOrders->phone, 
          'email'              => $objOrders->email,
          'count'              => $item['count'],
          'item_sku'           => $item['item_sku'],
          'item_name'          => $item['item_name'],
          'item_configuration' => $item['configuration'],
          'item_price'         => $item['item_price'],
          'item_sum'           => $item['sum'],
        );
      }         
    }
    
    // Output
    $this->saveToBrowser();
  }  
  
  /**
   * Export orders and send them to browser as file
   * @param DataContainer
   */
  public function exportBank()
  {    
    if ($this->Input->get('key') != 'export_bank') {
      return '';
    }

    $csvHead = &$GLOBALS['TL_LANG']['tl_iso_product_collection']['csv_head'];
    $arrKeys = array('company', 'lastname', 'firstname', 'street', 'postal', 'city', 'country', 'phone', 'email');
     
    foreach ($arrKeys as $v) {
      $this->arrHeaderFields[$v] = $csvHead[$v];
    }

    $objOrders = \Database::getInstance()->query("SELECT tl_iso_address.* FROM tl_iso_product_collection, tl_iso_address WHERE tl_iso_product_collection.billing_address_id = tl_iso_address.id AND ( document_number != '' AND document_number IS NOT NULL) GROUP BY member");

    if (null === $objOrders) {
      return '<div id="tl_buttons">
          <a href="'.ampersand(str_replace('&key=export_order', '', $this->Environment->request)).'" class="header_back" title="'.specialchars($GLOBALS['TL_LANG']['MSC']['backBT']).'">'.$GLOBALS['TL_LANG']['MSC']['backBT'].'</a>
          </div>
          <p class="tl_gerror">'. $GLOBALS['TL_LANG']['MSC']['noOrders'] .'</p>';
    }

    while ($objOrders->next()) {      
      $this->arrContent[$objOrders->id] = array(
        'company'       => $objOrders->company, 
        'lastname'      => $objOrders->lastname, 
        'firstname'     => $objOrders->firstname,
        'street'        => $objOrders->street_1, 
        'postal'        => $objOrders->postal, 
        'city'          => $objOrders->city, 
        'country'       => $GLOBALS['TL_LANG']['CNT'][$objOrders->country],
        'phone'         => $objOrders->phone, 
        'email'         => $objOrders->email,
      );       
    }
  
    // Output
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
    if (count($this->arrHeaderFields)>0) {
      $arrData = array($this->arrHeaderFields);
    }

    // add all other elements
    foreach ($this->arrContent as $k=>$v) {
      //TODO: maybe find a better solution
      $arrData[] = $v;
    }


    // build the csv string
    foreach((array) $arrData as $arrRow) {
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
