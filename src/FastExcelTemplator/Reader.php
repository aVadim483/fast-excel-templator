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


    public function __construct($file, ?array $parserProperties = [])
    {
        parent::__construct($file, $parserProperties);
    }

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
     * @return array
     */
    public function getAllAttributes(): array
    {

        return $this->nodes[$this->currentLevel][$this->currentNodeName]['__attr'] ?? [];
    }

    public function validate()
    {
        $this->setParserProperty(self::VALIDATE, true);
        foreach ($this->fileList() as $file) {
            echo $file, '<br>';
        }
    }
}

// EOF