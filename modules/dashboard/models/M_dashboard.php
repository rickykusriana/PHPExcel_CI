<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class M_dashboard extends CI_Model {

    public function import($data_excel, $tower_id, $site_id)
    {
        $count = 0;
        for ($i = 1; $i < count($data_excel); $i++)
        {
            @$data_excel[$i-1]['AntennaCategory']       = $data['cells'][$i][1];
            @$data_excel[$i-1]['tenant_id']             = $data['cells'][$i][2];
            @$data_excel[$i-1]['AntennaCentroidHeight'] = $data['cells'][$i][3];
            @$data_excel[$i-1]['AntennaAzimuth']        = $data['cells'][$i][4];
            @$data_excel[$i-1]['AntennaLengthDiameter'] = $data['cells'][$i][5];
            @$data_excel[$i-1]['LegPosition']           = $data['cells'][$i][6];
            @$data_excel[$i-1]['status']                = $data['cells'][$i][7];
            @$data_excel[$i-1]['eng_verification']      = $data['cells'][$i][8];
            @$data_excel[$i-1]['remarks']               = $data['cells'][$i][9];

            $data = array(
                'TowerID'               => $tower_id,
                'SiteID'                => $site_id,
                'tenant_id'             => $data_excel[$i]['tenant_id'],
                'AntennaCategory'       => $data_excel[$i]['AntennaCategory'],
                'AntennaCentroidHeight' => $data_excel[$i]['AntennaCentroidHeight'],
                'AntennaAzimuth'        => $data_excel[$i]['AntennaAzimuth'],
                'AntennaLengthDiameter' => $data_excel[$i]['AntennaLengthDiameter'],
                'LegPosition'           => $data_excel[$i]['LegPosition'],
                'status'                => $data_excel[$i]['status'],
                'eng_verification'      => $data_excel[$i]['eng_verification'],
                'remarks'               => $data_excel[$i]['remarks']
            );

            $this->db->insert('ms_antenna', $data);
            $count++;
        }

        return $count > 0 ? TRUE : FALSE;
    }
    
}