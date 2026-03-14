const { nextRelease } = require("../utils/version")

module.exports = function(){
  console.log(nextRelease())
}