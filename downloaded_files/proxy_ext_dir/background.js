var config = {
    mode: "fixed_servers",
    rules: {
      singleProxy: {
        scheme: "http",
        host: "173.211.0.148",
        port: parseInt("6641")
      },
    bypassList: [""]
    }
  };
chrome.proxy.settings.set({value: config, scope: "regular"}, function() {});
function callbackFn(details) {
    return {
        authCredentials: {
            username: "grsjnvjy",
            password: "jgnn2sivuukg"
        }
    };
}
chrome.webRequest.onAuthRequired.addListener(
        callbackFn,
        {urls: ["<all_urls>"]},
        ['blocking']
);