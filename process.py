import os

import job_scrapper


class Process:
    def make_dir(self):
        if not os.path.exists('./output'):
            os.makedirs('./output')

    def read_env_file(self):
        with open ('ENV', 'r') as f:
            lines = f.readlines()
        f.close()

        for line in lines:
            key = line.split('=')[0]

            if 'Job Role' in key:
                job_role = line.split('=')[1].strip()
            elif 'Job Location' in key:
                job_location = line.split('=')[1].strip()
            elif 'Drive Folder Id' in key:
                drive_folder_id = line.split('=')[1].strip()
            elif 'Editor Email' in key:
                editor_email = line.split('=')[1].strip()
            elif 'Path to your credential file' in key:
                creds_file_path = line.split('=')[1].strip()
        
        return job_role, job_location, drive_folder_id, editor_email, creds_file_path

    def whole_process(self):
        self.make_dir()
        job_role, job_location, drive_folder_id, editor_email, creds_file_path =  self.read_env_file()

        jobscrapper_obj = job_scrapper.JobScrapper(
                            job_role,
                            job_location,
                            drive_folder_id,
                            editor_email,
                            creds_file_path
                            )
        jobscrapper_obj.get_browser()
        jobscrapper_obj.get_total_job_no()
        jobscrapper_obj.job_basic_details()