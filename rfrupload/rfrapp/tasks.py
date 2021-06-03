from celery import shared_task
import time
from celery_progress.backend import ProgressRecorder
from .scripts.segmentation import sortFile, generateGroups
from .scripts.input import generateInput, updateReason

@shared_task(bind=True, name="logic", ignore_result=False)
def logic(self):
    progress_recorder = ProgressRecorder(self)
    print('Started')
    sortFile()
    progress_recorder.set_progress(1, 4, description="")
    groups = generateGroups()
    # print(groups['Discontinued/ Obselete'])
    # print(groups['Master data issue'])
    progress_recorder.set_progress(2, 4, description="")
    res = generateInput(groups)
    sales = res[0]
    material = res[1]
    progress_recorder.set_progress(3, 4, description="")
    updateReason(sales, material)
    progress_recorder.set_progress(4, 4, description="")
    return True