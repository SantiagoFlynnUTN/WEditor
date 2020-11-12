STRING="Configurando repositorio..."

echo $STRING

git checkout develop

git update-index --assume-unchanged WorldEditor.ini
git update-index --assume-unchanged WorldEditor.csi
git update-index --assume-unchanged WorldEditor.vbw
echo "WorldEditor.ini" >> .gitignore
echo "WorldEditor.csi" >> .gitignore  
echo "WorldEditor.vbw" >> .gitignore  
git rm --cached WorldEditor.ini
git rm --cached WorldEditor.csi
git rm --cached WorldEditor.vbw

cat .gitignore

echo "[CONFIGURACION]
GuardarConfig=1
UtilizarDeshacer=1
AutoCapturarTrans=1
AutoCapturarSup=0
ObjTranslado=0
[PATH]
UltimoMapa=
[MOSTRAR]
ControlAutomatico=1
Capa2=1
Capa3=1
Capa4=0
Translados=1
Objetos=1
NPCs=1
Triggers=0
Grilla=0
Decors=1
Bloqueos=0
LastPos=70-50
ClienteHeight=0 '  ancho en tiles (x32), full screen =0
ClienteWidth=0 '  alto en tiles (x32), full screen = 0
Capa9=0" > WorldEditor.ini

cat WorldEditor.ini