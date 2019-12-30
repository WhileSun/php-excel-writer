<?php
namespace ws\excel\writer;
use XLSXWriter;
/**
 * xlsxWriter导出
 * @author whilesun
 */
Class Export{

    private $width = 20; //col宽度
    private $DataType = 'string'; //数据类型
    private $writer; //XLSXWriter 插件
    private $col_keys ; //数组key

    public function __construct(){
        $this->writer = new XLSXWriter();
        $this->col_keys = array();
    }

    public function setConfig($config){
        $this->width = isset($config['width']) ? $config['width'] : $this->width;
        $this->DataType = isset($config['DataType']) ? $config['DataType'] : $this->DataType;
    }


    /**
     * 设置header 第一行数据
     * @param header_options Array header每列的属性
     * @param sheetName String sheet名称
     * @param options Array 配置全局属性 ['auto_filter'=>true,'freeze_rows'=>1,'freeze_columns'=>1]
     */
    public function setHeader($header_options,$sheetName="Sheet1",$col_options=array()){
        $header = array();
        $widths = array();
        //去重
        $this->col_keys[$sheetName] = array();
        foreach($header_options as $k => $v){
            $this->col_keys[$sheetName][] = $k;
            if(is_array($v)){
                $title = isset($v['title']) ? $v['title'] : $k.'_栏目';
                $type = isset($v['type']) ? $v['type'] : $this->DataType;
                $width = isset($v['width']) ? $v['width'] : $this->width;
                $header[$title] = $type;
                $widths[] =  $width;
            }else{
                $header[$v] = $this->DataType;
                $widths[] =  $this->width;
            }
        }
        $col_options['widths'] = $widths;
        $this->writer->writeSheetHeader($sheetName,$header,$col_options);
    }

    /**
     * 多条数据添加，数据量大的情况不推荐使用，可以使用setRow
     * @param rows 数据
     * @param sheetName sheet名称
     * @param row_options 个性参数
     * 主要使用参数
     * halign 对齐方向 left,right,center,none
     * color 字体颜色
     * fill 背景填充色
     * font-style 字体样式
     * font 字体类型
     * font-size 字体大小
     */
    public function setData($rows,$sheetName="Sheet1",$row_options=array()){
        $col_keys =  isset($this->col_keys[$sheetName]) ? $this->col_keys[$sheetName] : array();
        $col_keys_ret = empty($col_keys);
        foreach($rows as $row){
            $data = array();
            if($col_keys_ret){
                $data = $row;
            }else{
                foreach($col_keys as $key){
                    $data[$key] = isset($row[$key]) ? $row[$key] : '';
                }
            }
            $this->writer->writeSheetRow($sheetName,$data,$row_options);
            unset($data);
        }
    }


    /**
     * 单条数据添加
     * @param row 数据
     * @param sheetName sheet名称
     * @param row_options 个性参数
     * 主要使用参数
     * halign 对齐方向 left,right,center,none
     * color 字体颜色
     * fill 背景填充色
     * font-style 字体样式
     * font 字体类型
     * font-size 字体大小
     */
    public function setRow($row,$sheetName="Sheet1",$row_options=array()){
        $col_keys =  isset($this->col_keys[$sheetName]) ? $this->col_keys[$sheetName] : array();
        $data = array();
        if(empty($col_keys)){
            $data = $row;
        }else{
            foreach($col_keys as $key){
                $data[$key] = isset($row[$key]) ? $row[$key] : '';
            }
        }
        $this->writer->writeSheetRow($sheetName,$data,$row_options);
        unset($data);
    }

    /**
     * Excel导出或者保存
     * @param $filename 文件名
     * @param 导出|保存
     */
    public function load($filename,$isSave=false){
        if($isSave){
            $this->writer->writeToFile($filename);
        }else{
            ob_end_clean();
            header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
            header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            header('Content-Transfer-Encoding: binary');
            header('Cache-Control: must-revalidate');
            header('Pragma: public');
            $this->writer->writeToStdOut();
            exit(0);
        }
    }

    public function __destruct(){

    }
}
?>