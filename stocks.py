# Cria uma classe pra processar os dados posteriormente
class ativo():
    def __init__(self):
        self.nome = ""
        self.tipo = 0 # Ação ou fii
        self.classe = 0 # Perene ou tech - Tijolo ou papel (híbrido)
        self.cotacao = 0 # Valor atual da ação
        self.dy = 0 # Dividend yeld
        self.roe = 0 # ROE
        self.roa = 0 # ROA
        self.roic = 0 # ROIC
        self.pl = 0 # P/L
        self.pegr = 0 # PEG Ratio
        self.pvp = 0 # P/VP
        self.lpa = 0 # LPA
        self.vpa = 0 # VPA
        self.graham = (22.5*self.vpa*self.lpa)**(1/2) # Fórmula de Benjamin Graham
        self.divebitda = 0 # Div líquida/EBITDA
        self.divpl = 0 # Div líquida/PL
        self.lc = 0 # Líquidez corrente
        self.pa = 0 # Passivos/Ativos
        self.receita = 1 #	CAGR receitas
        self.lucro = 1 #	CAGR lucros
        self.relation = self.lucro/self.receita
        self.dyroa = self.dy - self.roa
        self.patrimonio = 0 # Valor patrimonial por cota
        self.dycagr = 0 # DY CAGR
        self.valorcagr = 0 # Valor CAGR
        # Geral
        self.nota_geral = []
        pass
    def cria_obj(self, tipo, indicadores=[], selic=12.0):
        for a in range(len(indicadores)):
            indicadores[a] = str(indicadores[a]).replace(",",".")
            indicadores[a] = str(indicadores[a]).replace("%","")
            if indicadores[a] == "-":
                indicadores[a] = "0.0"
        self.tipo = tipo 
        # Modelo FII
        if self.tipo == 2:
            self.nome = indicadores[0]
            self.cotacao = float(indicadores[1])
            self.dy = float(indicadores[4])
            self.patrimonio = float(indicadores[6])
            self.pvp = float(indicadores[7])
            self.dycagr = float(indicadores[9])
            self.valorcagr = float(indicadores[10])
            if self.dy >= selic:
                # Papel
                self.classe = 2
            if self.dy < selic:
                # Tijolo
                self.classe = 1
        # Modelo ação
        if self.tipo == 1:
            self.nome = indicadores[0]
            self.cotacao = float(indicadores[1])
            self.dy = float(indicadores[6])
            self.roe = float(indicadores[30])
            self.roa = float(indicadores[31])
            self.roic = float(indicadores[32])
            self.pl = float(indicadores[7])
            self.pegr = float(indicadores[8])
            self.pvp = float(indicadores[9])
            self.lpa = float(indicadores[16])
            self.vpa = float(indicadores[14])
            self.divebitda = float(indicadores[21])
            self.divpl = float(indicadores[20])
            self.lc = float(indicadores[25])
            self.pa = float(indicadores[24])
            self.receita = float(indicadores[34])
            self.lucro = float(indicadores[35])
            if self.graham < self.cotacao:
                #Setores perenes e empresas sólidas/ antigas
                self.classe = 1 
            else:
                # Small caps
                self.classe = 2
