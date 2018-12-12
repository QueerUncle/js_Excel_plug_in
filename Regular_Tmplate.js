/**
 *  2018/11/23  lize
 */
var Regular = {

  //判断字符长度
  gblen:function(value){

    var v = value;
    
      var len = 0;

    for (var i=0; i<v.length; i++) {

      if (v.charCodeAt(i)>127 || v.charCodeAt(i)==94) {

        len += 2;

      } else {

        len ++;

      }

    }

    return len;

  },
  //长度小于10
  len10:function(value){

    var l = this.gblen(value);

    if(l<10){

      return true;

    }else{

      return false;

    }

  },
  //不为为空
  notNull:function(value){

    var t = value.length;

    if(t == '' || t == null){

      return false;

    }else{

      return true;

    }

  },
  //是否是邮箱
  isEmail:function(value){

    var t = value;

    var reg = new RegExp("^[a-z0-9]+([._\\-]*[a-z0-9])*@([a-z0-9]+[-a-z0-9]*[a-z0-9]+.){1,63}[a-z0-9]+$");

    return reg.test(t);

  },
  //是否是移动电话
  isPhone:function(value){

    var t = value;

    var reg = new RegExp(/^1[34578]\d{9}$/);

    return reg.test(t);

  },
  //是否是固定电话
  isMob:function(value){

    var t = value;

    var reg = new RegExp(/^(\(\d{3,4}\)|\d{3,4}-|\s)?\d{7,14}$/);

    return reg.test(t);

  },
  //是否是身份证号
  isCard:function(value){

    var y = value;

    var reg = new RegExp(IDCard)

  },
  //判断里面是全英文
  isChinese:function(value){

    var t = value;

    var reg = new RegExp("^[a-zA-Z]+$");

    return reg.test(t);

  }
};

// export default Regular
