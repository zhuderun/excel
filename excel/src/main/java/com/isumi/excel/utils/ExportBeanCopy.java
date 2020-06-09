package com.isumi.excel.utils;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

public class ExportBeanCopy {

    public static void copyFromEntity(Object entity,Object exportBean){
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        if(entity!=null && exportBean != null){
            //获取实体的属性
            List<Field> fields = new ArrayList<>();
            Field[] entityFields = entity.getClass().getDeclaredFields();
            fields.addAll(Arrays.asList(entityFields));

            if(entity.getClass()!=null){
                Field[] superFields = entity.getClass().getSuperclass().getDeclaredFields();
                fields.addAll(Arrays.asList(superFields));
            }

            ArrayList<String> entityFieldNames = new ArrayList<>(entityFields.length);
            for(Field f:fields){
                entityFieldNames.add(f.getName());
            }

            Field[] resultFields = exportBean.getClass().getDeclaredFields();
            ArrayList<String> resultFieldNames = new ArrayList<>(entityFields.length);

            for(Field f:resultFields){
                resultFieldNames.add(f.getName());
            }

            for(String property:entityFieldNames){
                if(!"serialVersionUID".equals(property)){
                    try{
                        if(resultFieldNames.contains(property)){
                            Object fieldValue = Reflections.invokeGetter(entity,property);
                            if(fieldValue!=null){
                                if(fieldValue instanceof Date){
                                    Reflections.invokeSetter(exportBean,property,sdf.format(fieldValue));
                                }else{
                                    Reflections.invokeSetter(exportBean,property,String.valueOf(fieldValue));
                                }
                            }else{
                                Reflections.invokeSetter(exportBean,property,"");
                            }
                        }
                    }catch(Exception e){
                        e.printStackTrace();
                    }

                }
            }
        }
    }










}
