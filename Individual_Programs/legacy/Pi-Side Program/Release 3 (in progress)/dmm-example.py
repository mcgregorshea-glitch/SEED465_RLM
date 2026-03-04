from sys import exit
from signal import signal, SIGINT
from pyvisa import ResourceManager
from pyperclip import copy


DMM_CONFIG = [
    [120, 100, 'VINP'],
    [104, 100, 'IINP', 1e3],
    [107, 100, 'VSYS'],
    [103, 100, 'SAUX', 1e3],
    [102, 100, 'VINV'],
    [109, 100, 'SINV', 1e3],
]


class DmmInst:
    def __init__(self, id: int, samples: int, name: str, scale: float = 1) -> None:
        self.id = id
        self.samples = samples
        self.name = name
        self.scale = scale

    def connect(self, pvrmgr: ResourceManager) -> None:
        id_str = f'TCPIP0::10.123.210.{self.id}::inst0::INSTR'
        self.pv = pvrmgr.open_resource(id_str)

    def setup(self) -> None:
        self.pv.write(f'SAMP:COUN {self.samples}')
        self.pv.write(f'CALC:AVER:STAT ON')

    def trigger(self) -> None:
        self.pv.write('INIT')

    def ready(self) -> bool:
        return self.pv.query_ascii_values('CALC:AVER:COUN?')[0] >= self.samples

    def read(self) -> float:
        return self.pv.query_ascii_values('CALC:AVER:ALL?')[0] * self.scale


class DmmGroup:
    def __init__(self, config: list) -> None:
        self.dmms: list[DmmInst] = []
        for info in config:
            self.dmms.append(DmmInst(*info))

    def initialize(self) -> None:
        self.pvrmgr = ResourceManager()
        for dmm in self.dmms:
            dmm.connect(self.pvrmgr)
        for dmm in self.dmms:
            dmm.setup()

    def trigger(self) -> None:
        for dmm in self.dmms:
            dmm.trigger()

    def read(self) -> list[float]:
        ready = False
        while not ready:
            ready = True
            for dmm in self.dmms:
                ready &= dmm.ready()
        values: list[float] = []
        for dmm in self.dmms:
            values.append(dmm.read())
        return values


def exit_handler(*_):
    print()
    print(f'Exiting...')
    exit()


#########
# start #
#########

signal(SIGINT, exit_handler)

print(f'Initializing DMMs...')
dmms = DmmGroup(DMM_CONFIG)
dmms.initialize()

while True:
    print()
    input(f'Press Enter when ready...')

    print(f'Starting DMM sampling...')
    dmms.trigger()

    print(f'Reading DMM values...')
    values = dmms.read()

    copy('\t'.join(f'{x}' for x in values))
    print(f'Done! Results available on clipboard')
