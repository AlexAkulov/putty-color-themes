// Registry find library 
var Registry = function(iTree, sComputerName) {
  this.iTree = (iTree || Registry.HKEY_LOCAL_MACHINE);
  this.sComputerName = (sComputerName || '.');
  var locStr = 'winmgmts:{impersonationLevel=impersonate}//'
             + this.sComputerName
             + '/root/default:StdRegProv';
  this.oRegistry = GetObject(locStr);
};

// Add some useful constants to the Registry class.
(function() {
  var constants = { 'HKEY_CLASSES_ROOT':   0x80000000
                  , 'HKEY_CURRENT_USER':   0x80000001
                  , 'HKEY_LOCAL_MACHINE':  0x80000002
                  , 'HKEY_USERS':          0x80000003
                  , 'HKEY_CURRENT_CONFIG': 0x80000005
                  , 'HKEY_DYN_DATA':       0x80000006
                  , 'type': { 'REG_SZ':        1
                  ,           'REG_EXPAND_SZ': 2
                  ,           'REG_BINARY':    3
                  ,           'REG_DWORD':     4
                  ,           'REG_MULTI_SZ':  7 }
                  };
  for (var c in constants) { Registry[c] = constants[c]; }
})();

// Create Registry class methods which which wrap WMI calls.
(function () {
  var pro = Registry.prototype;
  var wmiSig = { 'EnumKey':                ['hDefKey', 'sSubKeyName']
               , 'EnumValues':             ['hDefKey', 'sSubKeyName']
               , 'GetBinaryValue':         ['hDefKey', 'sSubKeyName', 'sValueName']
               , 'GetDWORDValue':          ['hDefKey', 'sSubKeyName', 'sValueName']
               , 'GetExpandedStringValue': ['hDefKey', 'sSubKeyName', 'sValueName']
               , 'GetStringValue':         ['hDefKey', 'sSubKeyName', 'sValueName']
               , 'GetMultiStringValue':    ['hDefKey', 'sSubKeyName', 'sValueName']
               };
  for (var methodName in wmiSig) {
    pro[methodName] = (function(name, args) {
      return function(inSpec) {
        var oMeth = this.oRegistry.Methods_(name)
          , oIn = oMeth.inParameters.SpawnInstance_();
        for (var i in args) {
          if (inSpec[args[i]]) {
            oIn[args[i]] = inSpec[args[i]];
          }
        }
        return this.oRegistry.ExecMethod_(name, oIn);
      }
    })(methodName, wmiSig[methodName]);
  }
})();

// Traverses from the given root key and yields key/values to the handler.
Registry.prototype.find = function(rootKey, handler) {
  var StopWalkingException = function() { }
    , that = this;
  var findInner = function(key) {
    // Yield all values to the handler.
    var oOut = that.EnumValues({hDefKey: that.iTree, sSubKeyName:key});
    try {
      var aNames = oOut.sNames.toArray()
        , aTypes = oOut.Types.toArray();
      for (var i=0; i<aNames.length; i++) {
        // TODO: yield the appropriate value, not just strings.
        if (aTypes[i] === Registry.type.REG_SZ) {
          var valueName = key + '\\' + aNames[i];
          oOut = that.GetStringValue({hDefKey: that.iTree, sSubKeyName:key, sValueName:aNames[i]});
          if (!handler(valueName, oOut.sValue.toString())) {
            throw new StopWalkingException();
          }
        }
      }
    } catch (e) { // Ignore missing toArray method.
      if (e instanceof StopWalkingException) throw e;
    }
    // Recurse into any sub keys.
    oOut = that.EnumKey({hDefKey: that.iTree, sSubKeyName:key});
    try {
      var aSubKeys = oOut.sNames.toArray();
      for (var i=0; i<aSubKeys.length; i++) {
        findInner(key + '\\' + aSubKeys[i]);
      }
    } catch (e) { // Ignore missing toArray method.
      if (e instanceof StopWalkingException) throw e;
    }
  };
  // Search until naturally done or the handler says stop.
  try {
    findInner(rootKey);
  } catch (e) {
    if (!(e instanceof StopWalkingException)) { throw e; }
    return false;
  }
  return true;
};


// my code
var args = WScript.Arguments;
if (args.length<1) {
	WScript.Echo("Drag'n'Drop .reg file to this script in Windows Explorer");
	WScript.Quit(666);
}
var FilePath = args.item(0);

var result1 = WScript.CreateObject("WScript.Shell").Popup("Apply colors settings from "+FilePath+"?", 0, "Putty Color", 68);
if (result1!=6 ) { WScript.Quit(666); }

fs = new ActiveXObject("Scripting.FileSystemObject");
f = fs.GetFile(FilePath);
is = f.OpenAsTextStream( 1, 0 );
var Colours = [];
keyRegex = /"Colour(.*?)"$/

while( !is.AtEndOfStream ){
	var line = is.ReadLine();
	if (keyRegex.exec(line)) {
		var colar = line.split("=");
		Colours[colar[0].slice(1,-1)] = colar[1].slice(1,-1)
	}
}
is.Close();

var WSHShell = WScript.CreateObject("WScript.Shell");

var reg = new Registry(Registry.HKEY_CURRENT_USER)
      , rootKey = 'Software\\SimonTatham\\PuTTY\\Sessions'
      , keyRegex = /Sessions\\(.*?)\\Colour(.*?)$/
    reg.find(rootKey, function(path, value) {
      var keyMatch = path;
        if (keyRegex.exec(path)) {
          var m = keyMatch.split("\\");
		  var hashkey = m[m.length-1];
		  WSHShell.RegWrite("HKEY_CURRENT_USER\\"+path,Colours[hashkey],"REG_SZ");
        }
      return true; 
    });

WScript.Echo("Color setings from file "+FilePath+" were applied!");
