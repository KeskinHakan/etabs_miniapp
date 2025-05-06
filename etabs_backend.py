import comtypes.client
import pandas as pd
import os

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.float_format', lambda x: '%.2f' % x)
pd.options.mode.chained_assignment = None

class ETABSManager:
    def __init__(self, model_path):
        self.model_path = model_path
        self.start_etabs()
        self.open_model()

    def start_etabs(self):
        helper = comtypes.client.CreateObject('ETABSv1.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        self.EtabsObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
        self.EtabsObject.ApplicationStart()
        self.SapModel = self.EtabsObject.SapModel

    def start_safe(self):
        helper = comtypes.client.CreateObject('SAFEv1.Helper')
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        self.SafeObject = helper.CreateObjectProgID("CSI.SAFE.API.ETABSObject")
        self.SafeObject.ApplicationStart()
        self.SafeModel = self.SafeObject.SapModel

    def open_model(self):
        if not os.path.exists(self.model_path):
            raise FileNotFoundError(f"Model path does not exist: {self.model_path}")
        self.SapModel.File.OpenFile(self.model_path)
        self.SapModel.File.Save(self.model_path)
        self.SapModel.SetPresentUnits(6)

    def open_model_safe(self):
        if not os.path.exists(self.model_path):
            raise FileNotFoundError(f"Model path does not exist: {self.model_path}")
        self.SafeModel.File.OpenFile(self.model_path)
        self.SafeModel.File.Save(self.model_path)
        self.SafeModel.SetPresentUnits(6)

    def run_analysis(self):
        self.SapModel.Analyze.RunAnalysis()

    def run_analysis_safe(self):
        self.SafeModel.Analyze.RunAnalysis()

    def get_available_tables(self):
        ret, table_list, ret, ret, ret = self.SapModel.DatabaseTables.GetAvailableTables()
        return table_list

    def get_table_as_dataframe(self, table_key):
        ret, ret, column_names, record_number, data, ret = self.SapModel.DatabaseTables.GetTableForDisplayArray(table_key, ' ', "All")
        num_columns = len(column_names)
        rows = [data[i:i + num_columns] for i in range(0, len(data), num_columns)]
        df = pd.DataFrame(rows, columns=column_names)
        return df

    def process_first_last_station(self, df, id_columns):
        df["Station_float"] = df["Station"].str.replace(",", ".").astype(float)
        df = df.sort_values(id_columns + ["Station_float"])
        first_rows = df.groupby(id_columns, as_index=False).first()
        last_rows = df.groupby(id_columns, as_index=False).last()
        final_df = pd.concat([first_rows, last_rows], ignore_index=True)
        final_df = final_df.sort_values(id_columns + ["Station_float"]).reset_index(drop=True)
        return final_df

    def get_combo_list(self):
        ret, combo_list, ret = self.SapModel.RespCombo.GetNameList()
        ret, case_list, ret = self.SapModel.DatabaseTables.GetLoadCasesSelectedForDisplay()

        return combo_list, case_list

    def get_frame_list(self):
        ret, frame_list, _ = self.SapModel.FrameObj.GetNameList()

        return frame_list

    def get_frameforce_df(self):
        combo_list,case_list = self.get_combo_list()
        self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        for each in list(combo_list):
            self.SapModel.Results.Setup.SetComboSelectedForOutput(each)
        for each in list(case_list):
            self.SapModel.Results.Setup.SetCaseSelectedForOutput(each)

        frameforce_output = self.SapModel.Results.FrameForce("", 3)  # 3 = Global coordinates
        return self.frame_force_to_dataframe(frameforce_output)

    def frame_force_to_dataframe(self, frameforce_output):
        (
            number_of_results,
            frame_name,
            obj_station,
            elem_name,
            elem_station,
            output_case,
            step_type,
            step_num,
            P,
            V2,
            V3,
            T,
            M2,
            M3,
            ret_code
        ) = frameforce_output

        data = {
            "FrameName": frame_name,
            "ObjStation": obj_station,
            "ElemName": elem_name,
            "ElemStation": elem_station,
            "OutputCase": output_case,
            "StepType": step_type,
            "StepNum": step_num,
            "P": P,
            "V2": V2,
            "V3": V3,
            "T": T,
            "M2": M2,
            "M3": M3
        }

        df = pd.DataFrame(data)
        return df
    def close_etabs(self):
        self.EtabsObject.ApplicationExit(False)
