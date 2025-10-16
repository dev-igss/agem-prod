FROM php:8.2-apache

RUN apt-get update && apt-get install -y \
&& docker-php-ext-install mysqli pdo pdo_mysql

RUN apt-get update && apt-get install -y --no-install-recommends apt-utils libzip-dev libssl-dev \
    ## Zip extension
    && apt-get install -y --no-install-recommends zlib1g-dev \
    && docker-php-ext-install zip \
    ## Imap extension
    && apt-get install -y --no-install-recommends libc-client-dev libkrb5-dev \
    && PHP_OPENSSL=yes docker-php-ext-configure imap --with-kerberos --with-imap-ssl \
    && docker-php-ext-install imap \
    ## Xsl extension
    && apt-get install -y --no-install-recommends libxslt-dev \
    && docker-php-ext-install xsl \
    ## Intl extension
    && apt-get install -y --no-install-recommends libicu-dev g++ \
    && docker-php-ext-configure intl \
    && docker-php-ext-install intl \
    ## Mcrypt
    && apt-get install -y --no-install-recommends libmcrypt-dev \
    ## gd  & other image extentions
    && apt-get install -y --no-install-recommends libpng-dev libfreetype6-dev \
       libjpeg62-turbo-dev libpng-dev \
    && docker-php-ext-configure gd --with-freetype --with-jpeg \
    && docker-php-ext-install -j$(nproc) gd \
    ## Other extensions & database utils
    && apt-get install -y --no-install-recommends \
    && docker-php-ext-install mysqli pdo pdo_mysql \
    ## git, curl and unzip
    && apt-get install -y --no-install-recommends git curl unzip \
    ## Cron and spervisor
    && apt-get install -y --no-install-recommends cron rsyslog supervisor \
    && rm -rf /etc/cron.* \
    && mkdir /var/log/cron && touch /var/log/cron/cron.log \
    ## Clean-up
    && rm -rf /var/lib/apt/lists/*

RUN a2enmod rewrite

RUN echo "ServerName localhost" >> /etc/apache2/apache2.conf

RUN sed -i '/<Directory \/var\/www\/>/,/<\/Directory>/ s/AllowOverride None/AllowOverride All/' /etc/apache2/apache2.conf

RUN apt-get update && apt-get install -y tzdata
ENV TZ="America/Guatemala"

# Install Composer
RUN curl -sS https://getcomposer.org/installer | php -- --install-dir=/usr/local/bin --filename=composer

COPY ./code-dev /var/www

RUN chmod 777 -R -c /var/www

RUN chmod -R 777 /var/www/storage

RUN chmod -R 755 /var/www/storage

RUN chmod -R 755 /var/www/storage/logs

WORKDIR /var/www


