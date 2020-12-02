<template>
  <div class="hello">
    <h1>{{ msg }}</h1>
    <ul>

      <li @click="openDialog">6666</li>
      <li>
        <NotRelevance :title="title" :count="count" @on-click-carryRelevance="carryRelevance"></NotRelevance>
      </li>
    </ul>
    <AssociatedTaxpayer ref="child" :params="params" @selection-change='handleSelectionChange'></AssociatedTaxpayer>
    <UploadTemplateLink/>
  </div>
</template>

<script>
import AssociatedTaxpayer from "./AssociatedTaxpayer/index";
import UploadTemplateLink from './UploadTemplateLink/index'
export default {
   components: {AssociatedTaxpayer,UploadTemplateLink },
  props: {
    msg: String
  },
  data() {
    return {
      count:'75',
      title:'未关联纳税人的加油站个数',
      params:{
        dialogVisible:false,
        inputValue:'88',
        auxiliaryPayerId:'88888',
        titleName:''
      },
    }
    },
   methods: {
     carryRelevance(data){
       console.log('88',data)
     },
     openDialog(){
       this.params.dialogVisible=true
       this.params.titleName='888'
     },
     handleSelectionChange(selection){
        console.log(selection)
     },
  
   }
}
</script>

<style scoped>
h3 {
  margin: 40px 0 0;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>
