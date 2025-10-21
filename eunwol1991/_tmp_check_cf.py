import importlib.util
p = r"C:\Users\jhunj\puython\check data use\Check_File.py"
spec = importlib.util.spec_from_file_location('cf', p)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
print('short:', mod.shorten_path('A'+mod.os.sep+'B' +
      mod.os.sep+'C'+mod.os.sep+'D'+mod.os.sep+'E', 3))
print('valid:', mod.VALID_TAGS)
print('pat:', mod.FILENAME_PATTERN.pattern)
