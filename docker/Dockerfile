FROM openjdk:8-jre-alpine
RUN sed -i 's/dl-cdn.alpinelinux.org/mirrors.ustc.edu.cn/g' /etc/apk/repositories
RUN apk update && apk add tzdata && apk add ttf-dejavu && cp /usr/share/zoneinfo/Asia/Hong_Kong /etc/localtime \
&& echo "Asia/Shanghai" > /etc/timezone \
&& apk del tzdata