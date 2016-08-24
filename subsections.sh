#!/bin/sh
QUERY="(stereotactic OR stereotaxic) AND (radiotherapy OR radiosurgery) OR radiosurgery"
DATA=$1
python pyresearch_subsection.py -s "$QUERY AND (brain OR cranial OR cranium) AND (metastasis OR metastases)" -i $DATA -t "Brain Metastases"
python pyresearch_subsection.py -s "$QUERY AND (spine OR brainstem OR spinal cord) AND (metastasis OR metastases)" -i $DATA -t "Spinal Metastases"
python pyresearch_subsection.py -s "$QUERY AND Meningioma" -i $DATA -t "Meningioma"
python pyresearch_subsection.py -s "$QUERY AND (Glioblastoma OR GBM OR high grade glioma OR astrocytoma oligodendroglioma)" -i $DATA -t "Glioblastoma"
python pyresearch_subsection.py -s "$QUERY AND (Arteriovenous malformation OR AVM)" -i $DATA -t "AVM"
python pyresearch_subsection.py -s "$QUERY AND (Acoustic neuroma OR vestibular schwanoma)" -i $DATA -t "Acoustic Neuroma"
python pyresearch_subsection.py -s "$QUERY AND (Trigeminal Neuralgia OR Tic Douloureux)" -i $DATA -t "Trigeminal Neuralgia"
python pyresearch_subsection.py -s "$QUERY AND (depression OR anxiety disorder OR obsessive compulsive OR obsessive compulsion)" -i $DATA -t "Psychiatry"
