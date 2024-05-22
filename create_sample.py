from signal_processing.signals import read_signal, write_signal
import sys
import os

filename = os.path.expanduser(sys.argv[1])
export_filename = sys.argv[2]
signal = read_signal(filename)
signal.data = signal.data[:14336]
export_path = os.path.join('../', export_filename)
write_signal(export_path, signal)
