class TestPythonInstallationWorks():
    """docstring example this is a test class testing Python working properly,
   that's all folks"""
    def __init__(self):
        self.i = 32412
        self.j = "verily"
    def peer(self):
        self.i ** 2
        self.j += ", I say"
    def gear(self, num_in):
        self.i += num_in

a = TestPythonInstallationWorks()
a.peer()
print(str(a.i), a.j)
a.gear(23)
print(str(a.i))
print(a.__doc__)
for _ in range(0, 20):
    a.peer()
print(a.j)
