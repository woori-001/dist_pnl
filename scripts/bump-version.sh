#!/bin/sh
# 배포할 때마다 실행: 브라우저가 옛 파일과 새 파일을 섞어 쓰지 않도록 모든 참조에 같은 ?v= 를 붙인다
V=${1:-$(date +%Y%m%d%H%M)}
cd "$(dirname "$0")/.."
sed -i -E "s#(src/(app\.js|styles\.css))(\?v=[0-9a-z]+)?#\1?v=$V#g" index.html
sed -i -E "s#from '\./(config|rules|calc|eval)\.js(\?v=[0-9a-z]+)?'#from './\1.js?v=$V'#g" src/app.js src/calc.js src/eval.js
echo "version $V"
