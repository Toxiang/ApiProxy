yum install -y git curl ca-certificates openssl tar

docker build -t toxic/apiproxy:latest \
  --build-arg GOPROXY=https://goproxy.cn,direct \
  --build-arg GOSUMDB=sum.golang.google.cn \
  -f Dockerfile .