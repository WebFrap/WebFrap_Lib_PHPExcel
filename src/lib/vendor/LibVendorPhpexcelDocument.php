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


/**
 * @package Webfrap
 * @subpackage ModBase
 * @author Dominik Bonsch <dominik.bonsch@webfrap.de>
 * @copyright Webfrap Developer Network <contact@webfrap.de>
 */
class LibVendorPhpexcelDocument
  extends PHPExcel
{

	/**
	 * Create a new PHPExcel with one Worksheet
	 * @param string $title
	 * @param string $sheetClass
	 */
	public function __construct( $title = null, $sheetClass = null )
	{
	  
	  parent::__construct();
	  
		// Initialise worksheet collection and add one worksheet
		$this->_workSheetCollection = array();
		
		if( $sheetClass && Webfrap::classLoadable( $sheetClass ) )
		{  
		  $this->_workSheetCollection[0] = new $sheetClass( $this, $title );
		}
		else
		{ 
		  $this->_workSheetCollection[0] = new PHPExcel_Worksheet( $this, $title );
		}
		  
		$this->_activeSheetIndex = 0;
    
		/*
		// Create document properties
		$this->_properties = new PHPExcel_DocumentProperties();

		// Create document security
		$this->_security = new PHPExcel_DocumentSecurity();

		// Set named ranges
		$this->_namedRanges = array();

		// Create the cellXf supervisor
		$this->_cellXfSupervisor = new PHPExcel_Style(true);
		$this->_cellXfSupervisor->bindParent($this);

		// Create the default style
		$this->addCellXf(new PHPExcel_Style);
		$this->addCellStyleXf(new PHPExcel_Style);
		*/
		
	}//end public function __construct */
	
    /**
     * Create sheet and add it to this workbook
     *
	 * @param int|null $iSheetIndex Index where sheet should go (0,1,..., or null for last)
     * @return PHPExcel_Worksheet
     * @throws Exception
     */
    public function createSheet($iSheetIndex = NULL, $className = null, $title = null )
    {
      
      if( $className )
      {
        $newSheet = new $className( $this, $title );
      }
      else 
      {
        $newSheet = new PHPExcel_Worksheet($this);
      }
      
      $this->addSheet($newSheet, $iSheetIndex);
      return $newSheet;
    }
    
	/**
	 * Get sheet by index
	 *
	 * @param int $pIndex Sheet index
	 * @return PHPExcel_Worksheet
	 * @throws Exception
	 */
	public function getSheet($pIndex = 0)
	{
		$numSheets = count($this->_workSheetCollection) - 1;

		if ( $pIndex > $numSheets ) 
		{
		  
			throw new InternalError_Exception
			(
				"Sheet index: {$pIndex} is out of bounds. There are only {$numSheets} sheets."
			);
		} 
		else 
		{
			return $this->_workSheetCollection[$pIndex];
		}
	}
  

}//end class LibVendorPhpexcelDocument




