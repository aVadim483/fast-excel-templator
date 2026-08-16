<?php

namespace avadim\FastExcelTemplator;

/**
 * Class Reader
 *
 * @package avadim\FastExcelReader
 */
class Reader extends \avadim\FastExcelReader\Reader
{
    protected string $zipFile;

    protected ?string $innerFile = null;

    protected array $xmlParserProperties = [];

    protected array $nodes = [];

    protected int $currentLevel = -1;

    protected string $currentNodeName = '';


    /**
     * Reader constructor
     *
     * @param string $file
     * @param array|null $parserProperties
     */
    public function __construct($file, ?array $parserProperties = [])
    {
        parent::__construct($file, $parserProperties);
    }

    /**
     * Read next node
     *
     * @return bool
     */
    public function read(): bool
    {
        $result = parent::read();

        if ($this->nodeType === \XMLReader::ELEMENT) {
            $this->currentLevel = $this->depth;
            $this->currentNodeName = $this->name;
            $this->nodes[$this->currentLevel][$this->currentNodeName]['__attr'] = [];
            if ($this->hasAttributes) {
                while ($this->moveToNextAttribute()) {
                    $this->nodes[$this->currentLevel][$this->currentNodeName]['__attr'][$this->name] = $this->value;
                }
                $this->moveToElement();
            }
        }

        return $result;
    }

    /**
     * Get attributes of specified node
     *
     * @param string $nodeName
     *
     * @return array|null
     */
    public function getNodeAttributes(string $nodeName): ?array
    {
        for ($level = $this->currentLevel; $level >= 0; $level--) {
            if (isset($this->nodes[$level][$nodeName]['__attr'])) {
                return $this->nodes[$level][$nodeName]['__attr'];
            }
        }

        return null;
    }

    /**
     * Get all attributes of current node
     *
     * @return array
     */
    public function getAllAttributes(): array
    {

        return $this->nodes[$this->currentLevel][$this->currentNodeName]['__attr'] ?? [];
    }

    /**
     * List the inner files of the template
     *
     * It used to echo them with <br> tags, which a library has no business doing — print the
     * returned list yourself if you need it.
     *
     * Parser validation is switched on only when a document is already open: XMLReader refuses
     * to touch parser properties before that, and the file list does not depend on them, so the
     * method answers instead of failing when called right after the reader was created.
     *
     * @return array List of files inside the template
     */
    public function validate()
    {
        if ($this->innerFile !== null) {
            $this->setParserProperty(self::VALIDATE, true);
        }

        return $this->fileList();
    }
}

// EOF