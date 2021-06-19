<?php


abstract class msword
{
    static $msword = false;
    static $version;
    static $document = false;
    

    //////////////////////////////////////////////////////////////////////////////////////////
    
    static function open($word_filename)
    {
        self::open_application();
        $word_filename = str_replace('/', '\\', $word_filename);
        // if Open function returns NULL:
        // Open DCOM Config Settings:
        // Start -> dcomcnfg.exe
        // Computer
        // Local Computer
        // Config DCOM
        // Search for Microsoft Word 97-2003 Documents -> Properties
        // Tab Identity, change from Launching User to Interactive User
        self::$document = self::$msword->Documents->Open($word_filename);
        if(!self::$document) self::$document = false;
        else self::$document->Activate();
    } 

    static function close()
    {
        if((!self::$msword) || (!self::$document)) return;
        self::$document->Close(false);
        self::$document = false;
    }     

    static function save($new_word_filename, $format = 'docx')
    {
        if((!self::$msword) || (!self::$document))
            throw new Exception('New or opend document is required before this action.');

        if(file_exists($new_word_filename)) unlink($new_word_filename); // delete file if exists
        $format_list = [
            'doc'   => 0,     // Microsoft Word 97 document format
            'docx'  => 16,    // For Microsoft Office Word 2007, this is the DOCX format.
            'pdf'   => 17     // PDF format          
        ];
        $format = strtolower($format);
        $format_code = key_exists($format, $format_list) ? $format_list[$format] : 16;

        //https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._document.saveas?view=word-pia
        //https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdsaveformat?view=word-pia
        self::$document->SaveAs(
            $new_word_filename, 
            new VARIANT($format_code, VT_I4)
        ); // save the new file
    }

    // https://docs.microsoft.com/en-us/office/vba/api/word.find.execute
    static function find($findText, $options = [])
    {
        if((!self::$msword) || (!self::$document))
        throw new Exception('New or opend document is required before this action.');

        return self::$msword->Selection->Find->Execute(
            $findText,
            key_exists('matchCase', $options) ? $options['matchCase'] : false,
            key_exists('matchWholeWord', $options) ? $options['matchWholeWord'] : true,
            key_exists('matchWildCards', $options) ? $options['matchWildCards'] : false,
            key_exists('matchSoundsLike', $options) ? $options['matchSoundsLike'] : false,
            key_exists('matchAllWordForms', $options) ? $options['matchAllWordForms'] : false,
            key_exists('forward', $options) ? $options['forward'] : true,
            key_exists('wrap', $options) ? $options['wrap'] : 1,
            key_exists('format', $options) ? $options['format'] : false,
            key_exists('replaceWithText', $options) ? $options['replaceWithText'] : '',
            key_exists('replace', $options) ? $options['replace'] : 0,
            key_exists('matchKashida', $options) ? $options['matchKashida'] : false,
            key_exists('matchDiacritics', $options) ? $options['matchDiacritics'] : false,
            key_exists('matchAlefHamza', $options) ? $options['matchAlefHamza'] : false,
            key_exists('matchControl', $options) ? $options['matchControl'] : false
        );      
    }   
    
    static function text()
    {
        return self::$msword->Selection->Text;
    }

    static function select()
    {
        self::$document->Select();
    }

    // wdAllowOnlyFormFields	2	Allow content to be added to the document only through form fields.
    // wdAllowOnlyReading	3	Allow read-only access to the document.
    // wdAllowOnlyRevisions	0	Allow only revisions to be made to existing content.
    // wdNoProtection	-1	Do not apply protection to the document.
    // https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.tools.word.document.protect?view=vsto-2017         
    static function protect($WdProtectionType, $Password = '', $NoReset = false, $UseIRM = false, $EnforceStyleLock = true)
    {
        if((!self::$msword) || (!self::$document))
        throw new Exception('New or opend document is required before this action.');

        // help to findout functions
        //com_print_typeinfo(self::$document);
        //com_print_typeinfo(self::$msword);
        self::$document->Protect(
            new VARIANT($WdProtectionType, VT_I4),
            $NoReset,
            $Password,
            $UseIRM,
            $EnforceStyleLock
        );                   
    }

    static function open_application()
    {
        if(self::$msword) return;   // return if msword application opend before
      
        try
        {
            // dot net com must be enable
            // this 2 line must be added to php.ini
            // [COM_DOT_NET]
            // extension=php_com_dotnet.dll
            // get last opened msword application
            self::$msword = com_get_active_object('word.application');
        }
        catch(Exception $e)
        {
            // open new msword application
            self::$msword = new COM('word.application', null, CP_UTF8);      
        }
        if(!self::$msword) throw new Exception('Microsoft Word application must be installed on the server.');
        self::$version = self::$msword->Version;
        self::$msword->Visible = 1; //bring it to front
        self::$msword->DisplayAlerts = 0;
    }

    static function close_application()
    {
        if(!self::$msword) return;           // return if msword application closed before
        if(self::$document) self::close();   // close document if exists
        self::$document = false;

        //closing word
        self::$msword->Quit();
        self::$msword = false;
    }

    //////////////////////////////////////////////////////////////////////////////////////////
}

