<?php
defined('BASEPATH') OR exit('No direct script access allowed');

require APPPATH . 'libraries/PHPExcel.php';

class Dashboard extends MX_Controller {
    
    public function __construct()
    {
        parent::__construct();
        $this->load->model('m_dashboard');
    }

    public function import_excel()
    {
        if (isset($_POST['submit'])) {

            $site_id    = $this->input->post('site_id');
            $tower_id   = $this->input->post('tower_id');

            if (empty($_FILES['file']['name'])) {

                $this->session->set_flashdata('alert-danger', 'File belum dipilih');
                redirect(site_url('sites/detail/'.$site_id).'#Tower');
            }
            else {

                // Load helper
                $this->load->helper('file');

                $config['upload_path']   = './assets/file/';
                $config['allowed_types'] = 'xls';

                // Load library
                $this->load->library('upload', $config);
                $this->upload->initialize($config);

                if ( ! $this->upload->do_upload('file')) {

                    $this->session->set_flashdata('alert-danger', $this->upload->display_errors());
                    redirect(site_url('sites/detail/'.$site_id).'#Tower');
                }
                else {

                    $upload_data = $this->upload->data();

                    // Load library
                    $this->load->library('excel_reader');
                    $this->excel_reader->setOutputEncoding('CP1251');
                    $this->excel_reader->read($upload_data['full_path']);

                    $data       = $this->excel_reader->sheets[0];
                    $data_excel = array();

                    for ($i=1; $i<=$data['numRows']; $i++) {

                        if ($data['cells'][$i][1] == '') break;

                        @$data_excel[$i-1]['AntennaCategory']       = $data['cells'][$i][1];
                        @$data_excel[$i-1]['tenant_id']             = $data['cells'][$i][2];
                        @$data_excel[$i-1]['AntennaCentroidHeight'] = $data['cells'][$i][3];
                        @$data_excel[$i-1]['AntennaAzimuth']        = $data['cells'][$i][4];
                        @$data_excel[$i-1]['AntennaLengthDiameter'] = $data['cells'][$i][5];
                        @$data_excel[$i-1]['LegPosition']           = $data['cells'][$i][6];
                        @$data_excel[$i-1]['status']                = $data['cells'][$i][7];
                        @$data_excel[$i-1]['eng_verification']      = $data['cells'][$i][8];
                        @$data_excel[$i-1]['remarks']               = $data['cells'][$i][9];
                    }

                    @unlink('./assets/file/'.$upload_data['file_name']);

                    // Access model
                    if ($this->m_dashboard->import($data_excel, $tower_id, $site_id)) {

                        $this->session->set_flashdata('alert-success', 'Impor data berhasil');
                        redirect(site_url('sites/detail/'.$site_id).'#Tower');
                    }
                    else {

                        $this->session->set_flashdata('alert-danger', 'Gagal impor data, cek kembali');
                        redirect(site_url('sites/detail/'.$site_id).'#Tower');
                    }
                }
            }
        }
        else {
            show_404();
        }
    }

    public function xls_template()
    {
        $filename   = 'TEMPLATE-ANTENNA';
        $objXLS     = new PHPExcel();
        $no         = 1;
        $font       = array('font' => array('bold' => true));
        $styleArray = array(
                          'borders' => array(
                                       'allborders' => array(
                                                       'style' => PHPExcel_Style_Border::BORDER_THIN,
                                                       'color' => array(
                                                                  'rgb'  => '000000'
                                                                ),
                                                       ),
                                       ),
                            );

        $objBarang = $objXLS->setActiveSheetIndex(0);
        $objBarang->setCellValue('A1', 'ANTENNA CATEGORY');
        $objBarang->setCellValue('B1', 'TENANT NAME');
        $objBarang->setCellValue('C1', 'HEIGHT (m)');
        $objBarang->setCellValue('D1', 'AZIMUTH');
        $objBarang->setCellValue('E1', 'LENGTH (m)');
        $objBarang->setCellValue('F1', 'LEG POS');
        $objBarang->setCellValue('G1', 'STATUS');
        $objBarang->setCellValue('H1', 'VERIFICATION');
        $objBarang->setCellValue('I1', 'REMARK');

        // Result object from database
        foreach ($this->config_m->get_antenna_category() as $key) {
            $category[] = $key->name;
        }

        // Result object from database
        foreach ($this->config_m->get_tenant() as $key) {
            $tenant[] = $key->code_operator;
        }
        
        // Convert result array to string from result
        $list_category  = str_replace(array('"', '[', ']'), array('', '', ''), json_encode($category));
        $list_tenant    = str_replace(array('"', '[', ']'), array('', '', ''), json_encode($tenant));

        $list_legPos = 'A, B, C, D';
        $list_status = 'Plan by CAF, Plan by COLO, Existing by CAF, Remove by CAF, Dismantle';
        $list_verification = 'Verified, Not Verified';

        // Match to cell alphabet in Excel
        $data = array(
                    'A' => $list_category, 
                    'B' => $list_tenant,
                    'F' => $list_legPos,
                    'G' => $list_status,
                    'H' => $list_verification );

        foreach ($data as $key => $value) {

            // Show 100 rows in result excel
            for ($i=2; $i < 100; $i++) { 
                $objValidation = $objXLS->getActiveSheet()->getCell($key.$i)->getDataValidation();
                $objValidation->setType( PHPExcel_Cell_DataValidation::TYPE_LIST );
                $objValidation->setErrorStyle( PHPExcel_Cell_DataValidation::STYLE_INFORMATION );
                $objValidation->setAllowBlank(false);
                $objValidation->setShowInputMessage(true);
                $objValidation->setShowErrorMessage(true);
                $objValidation->setShowDropDown(true);
                $objValidation->setErrorTitle('Input error');
                $objValidation->setError('Value is not in list');
                $objValidation->setPromptTitle('- Choose One -');
                $objValidation->setPrompt('Please pick a value from the dropdown list');
                $objValidation->setFormula1('"'.$value.'"');
            }
        }

        foreach(range('A', 'I') as $alphabet) {
            $objXLS->getActiveSheet()->getColumnDimension($alphabet)->setAutoSize(true);
        }

        $objXLS->getActiveSheet()
               ->getStyle('A1:I1')
               ->applyFromArray($font);

        $objXLS->setActiveSheetIndex(0);        
        $objBarang->getStyle('A1:I'.$no)->applyFromArray($styleArray);

        $objWriter = PHPExcel_IOFactory::createWriter($objXLS, 'Excel5'); 
        header('Content-Type: application/vnd.ms-excel'); 
        header('Content-Disposition: attachment;filename="'.$filename.'.xls"'); 
        header('Cache-Control: max-age=0'); 
        $objWriter->save('php://output'); 
        exit();
    }

}
