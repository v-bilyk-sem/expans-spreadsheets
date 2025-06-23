function getStatusCode(url) {
  var url_trimmed = url.trim();
  if (url_trimmed === '') {
    return '';
  }

  var cache = CacheService.getScriptCache();
  var cacheKey = 'status-' + url_trimmed;
  var result = cache.get(cacheKey);

  if (!result) {
    var options = {
      'muteHttpExceptions': true,
      'followRedirects': false
    };

    try {
      var response = UrlFetchApp.fetch(url_trimmed, options);
      var responseCode = response.getResponseCode();
      result = responseCode.toString();
      cache.put(cacheKey, result, 21600); // Store in cache for 6 hours
    } catch (error) {
      result = 'Error: Unable to fetch URL';
    }
  }

  return result;
}
