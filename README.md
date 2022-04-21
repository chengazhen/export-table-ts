[使用的demo](https://gitee.com/a_zhen_a/export-excel)

## 使用示例(支持多级 Header)

```javascript
import {
    exportData
} from '@azhena/exportexcel'

/**
 * @description:
 * @param {*} list
 * @param {*} revealList 表头对应的数据属性
 * @param {*} 测试 文件名
 * @return {*}
 */
exportData(list, revealList, '测试') //

// 

const list = [{
        name: '张三',
        js: '熟练',
        css: '一般',
        nio: '了解',
        basic: '精通',
        springboot: '熟练',
        mybatis: '了解'
    },
    {
        name: '张三',
        js: '熟练',
        css: '一般',
        nio: '了解',
        basic: '精通',
        springboot: '熟练',
        mybatis: '了解'
    },
    {
        name: '张三',
        js: '熟练',
        css: '一般',
        nio: '了解',
        basic: '精通',
        springboot: '熟练',
        mybatis: '了解'
    },
    {
        name: '张三',
        js: '熟练',
        css: '一般',
        nio: '了解',
        basic: '精通',
        springboot: '熟练',
        mybatis: '了解'
    }
]
const revealList = [{
        name: '姓名',
        prop: 'name'
    },
    {
        name: '专业技能',
        child: [{
                name: '前端',
                child: [{
                        name: 'JavaScript',
                        prop: 'js'
                    },
                    {
                        name: 'CSS',
                        prop: 'css'
                    }
                ]
            },
            {
                name: '后端',
                child: [{
                        name: 'java',
                        child: [{
                                name: 'nio',
                                prop: 'nio'
                            },
                            {
                                name: '基础',
                                prop: 'basic'
                            }
                        ]
                    },
                    {
                        name: '框架',
                        child: [{
                                name: 'SpringBoot',
                                prop: 'springboot'
                            },
                            {
                                name: 'MyBatis',
                                prop: 'mybatis'
                            }
                        ]
                    }
                ]
            }
        ]
    }
]
```
