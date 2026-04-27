yum install -y git curl ca-certificates openssl tar

docker build -t toxic/apiproxy:latest \
  --build-arg GOPROXY=https://goproxy.cn,direct \
  --build-arg GOSUMDB=sum.golang.google.cn \
  -f Dockerfile .

docker run -d   --name apiproxy --restart unless-stopped  \
          --network api-proxy-net  -p 18080:8080    \
          -v /home/ai/data/proxy:/app/data   \
          --env-file  /home/ai/ApiProxy/deploy/.env   \
          toxic/apiproxy:latest