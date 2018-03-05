<?php
/**
 * project: iningmeng.
 * Author: iningmeng
 * Email: iningmeng@qq.com
 * Date: 2018/3/5  10:15
 * Version: 1.0
 * Description:
 */
namespace iningmeng\pptimages;
class Handle
{
    public function index($file_path,$out_path){
        $powerpnt = new \COM("powerpoint.application") or die("Unable to instantiate Powerpoint");
        if (!file_exists($file_path)){
            return 'file not exit';
        }
        echo realpath($file_path)."<br/>";
        $file_path = realpath($file_path);
        $out_path = realpath($out_path);
        try{
            $presentation = $powerpnt->Presentations->Open($file_path, false, false, false) or die("Unable to open presentation");
//            $presentation->Fonts->Replace('黑体','幼圆');
//            $presentation->Fonts->Replace('MS Gothic','幼圆');
//            $presentation->Fonts->Replace('方正粗倩简体','幼圆');
//            $presentation->Fonts->Replace('方正小标宋简体','幼圆');
//            $presentation->Fonts->Replace('Arial Black','幼圆');
//            $presentation->Fonts->Replace('华文中宋','幼圆');
//            $presentation->Fonts->Replace('Arial Unicode MS','幼圆');
//            $presentation->Fonts->Replace('方正细圆简体','幼圆');
//            $presentation->Fonts->Replace('Times New Roman','幼圆');

            foreach($presentation->Slides as $slide)
            {
                $slideName = "Slide_" . $slide->SlideNumber;
                $exportFolder = realpath($out_path);
                $slide->Export($exportFolder."//".$slideName.".jpg", "jpg", "1920", "1080");
            }
            $presentation->Close();
            $powerpnt->Quit();
            $powerpnt = null;
        }catch (\Exception $e){
            echo $e;
        }

    }
}