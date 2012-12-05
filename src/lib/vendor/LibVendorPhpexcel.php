<?php
/*******************************************************************************
*
* @author      : Dominik Bonsch <dominik.bonsch@webfrap.net>
* @date        :
* @copyright   : Webfrap Developer Network <contact@webfrap.net>
* @project     : Webfrap Web Frame Application
* @projectUrl  : http://webfrap.net
*
* @licence     : BSD License see: LICENCE/BSD Licence.txt
* 
* @version: @package_version@  Revision: @package_revision@
*
* Changes:
*
*******************************************************************************/


if( !defined( 'VENDOR_PHPEXCEL_VERSION' ) )
{

  define( 'VENDOR_PHPEXCEL_VERSION' , 'actual' );

  // include the needed data
  include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel.php';
  
  /** PHPExcel_Cell_AdvancedValueBinder */
  include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Cell/AdvancedValueBinder.php';
  
  /** PHPExcel_IOFactory */
  include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/IOFactory.php';
  
  /** Read Filters */
  include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Reader/IReadFilter.php';

}

// Set value binder
PHPExcel_Cell::setValueBinder( new PHPExcel_Cell_AdvancedValueBinder() );

/**
 * @package Webfrap
 * @subpackage ModBase
 * @author Dominik Bonsch <dominik.bonsch@webfrap.de>
 * @copyright Webfrap Developer Network <contact@webfrap.de>
 * @licence GPLv3
 */
class LibVendorPhpexcel
{

  /**
   * simple method just to call the autoload
   *
   */
  public static function init()
  {
  }//end public static function init */
  

/*//////////////////////////////////////////////////////////////////////////////
// Readers
//////////////////////////////////////////////////////////////////////////////*/

  /**
   * 
   */
  public static function loadReader2007()
  {
    // include the needed data
    include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Reader/Excel2007.php';

  }//end public static function loadReader2007 */

  /**
   * 
   */
  public static function loadReader2003()
  {
    // include the needed data
    include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Reader/Excel2003XML.php';

  }//end public static function loadReader2003 */

  /**
   * 
   */
  public static function loadReader5()
  {
    // include the needed data
    include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Reader/Excel5.php';

  }//end public static function loadReader5 */

  /**
   * 
   */
  public static function loadReaderOOCalc()
  {
    // include the needed data
    include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Reader/OOCalc.php';
  }//end public static function loadReaderOOCalc */

/*//////////////////////////////////////////////////////////////////////////////
// Writers
//////////////////////////////////////////////////////////////////////////////*/
  
  /**
   * 
   */
  public static function loadWriter2007()
  {
    // include the needed data
    include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Writer/Excel2007.php';
  }//end public static function loadWriter2007 */
  
/*//////////////////////////////////////////////////////////////////////////////
// Some Loaders
//////////////////////////////////////////////////////////////////////////////*/
  
  /**
   * 
   * @param PHPExcel $phpExcel
   */
  public static function getExcelWriter2007( $phpExcel )
  {
    
    if( !class_exists('PHPExcel_Writer_Excel2007') )
      self::loadWriter2007();
      
    return PHPExcel_IOFactory::createWriter($phpExcel, "Excel2007");
    
  }//end public static function getExcelWriter2007 */
  
  /**
   * 
   */
  public static function newSheet()
  {
    return new PHPExcel();
  }//end public static function newSheet  */
  
  
  
  public static function loadSheet($path)
  {
  	
  	$objReader = new PHPExcel_Reader_Excel2007();
	  $objPHPExcel = $objReader->load($path);
  	
    return $objPHPExcel;
  }//end public static function newSheet  */


}//end class LibVendorPhpexcel




