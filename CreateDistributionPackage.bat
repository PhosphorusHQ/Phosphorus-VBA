del PhosphorusPackage.tar
del PhosphorusPackage.tar.SHA256.txt

tar -caf PhosphorusPackage.tar ".\src"
tar -rf PhosphorusPackage.tar Phosphorus.xlam pPath.xlam pUnit.xlam Phosphorus Externals.xlam PhosphorEssence.xlam 
tar -rf PhosphorusPackage.tar "Phosphorus Tests.xlam" "PhosphorEssence - Examples.xlam" "pUnit Tests - pUnit.xlam"  "pUnit Tests - pPath.xlam" 
tar -rf PhosphorusPackage.tar "pUnit Test Runs - pUnit.xlam" "pUnit Test Runs - pPath.xlam"
tar -rf PhosphorusPackage.tar ".\images"
tar -rf PhosphorusPackage.tar ".\Tests"
tar -rf PhosphorusPackage.tar LogReader.xlsx
tar -rf PhosphorusPackage.tar LICENSE

certutil -hashfile PhosphorusPackage.tar SHA256 > PhosphorusPackage.tar.SHA256.txt
