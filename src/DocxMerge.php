<?php
/**
 * User: krustnic
 * Date: 04.02.14
 * Time: 11:17
 */

class DocxMerge {

    public function __construct()
    {
        //Make sure autoloader is loaded
        if (version_compare(PHP_VERSION, '5.1.2', '>=') and
            !spl_autoload_functions() || !in_array('DocxMergeAutoload', spl_autoload_functions())) {
            require dirname(__FILE__).DIRECTORY_SEPARATOR.'DocxMerge'.DIRECTORY_SEPARATOR.'DocxMergeAutoload.php';
        }
    }

    /**
     * Merge files in $docxFilesArray order and
     * create new file $outDocxFilePath
     * @param $docxFilesArray
     * @param $outDocxFilePath
     * @return int
     */
    public function merge( $docxFilesArray, $outDocxFilePath) {
        if ( count($docxFilesArray) == 0 || count($docxFilesArray) == 1 ) {
            // No files to merge
            return -1;
        }

        $toMerge = [
            $docxFilesArray[0],
            $docxFilesArray[1]
        ];

        $zip = new clsTbsZip();

        // Open the first document
        $zip->Open($toMerge[0]);
        $content1 = $zip->FileRead('word/document.xml');
        $zip->Close();

        // Extract the content of the first document
        $p = strpos($content1, '<w:body');
        if ($p===false) exit("Tag <w:body> not found in document 1.");
        $p = strpos($content1, '>', $p);
        $content1 = substr($content1, $p+1);
        $p = strpos($content1, '</w:body>');
        if ($p===false) exit("Tag </w:body> not found in document 1.");
        $content1 = substr($content1, 0, $p);

        // Insert into the second document
        $zip->Open($toMerge[1]);
        $content2 = $zip->FileRead('word/document.xml');
        $p = strpos($content2, '</w:body>');
        if ($p===false) exit("Tag </w:body> not found in document 2.");
        $content2 = substr_replace($content2, $content1, $p, 0);
        $zip->FileReplace('word/document.xml', $content2, TBSZIP_STRING);

        $continueMerge = [
            $toMerge[1] . '.docx'
        ];
        foreach($docxFilesArray as $key => $file) {
            if($key <= 1) continue;
            $continueMerge[] = $file;
        }

        if(count($continueMerge) > 1) {
            $zip->Flush(TBSZIP_FILE, $toMerge[1] . '.docx');

            return $this->merge($continueMerge, $outDocxFilePath);
        }

        $zip->Flush(TBSZIP_FILE, $outDocxFilePath);
        return 0;
    }

    public function setValues( $templateFilePath, $outputFilePath, $data ) {

        if ( !file_exists( $templateFilePath ) ) {
            return -1;
        }

        if ( !copy( $templateFilePath, $outputFilePath ) ) {
            // Cannot create output file
            return -2;
        }

        $docx = new Docx( $outputFilePath );
        $docx->loadHeadersAndFooters();
        foreach( $data as $key => $value ) {
            $docx->findAndReplace( "\${".$key."}", $value );
        }

        $docx->flush();
    }

}