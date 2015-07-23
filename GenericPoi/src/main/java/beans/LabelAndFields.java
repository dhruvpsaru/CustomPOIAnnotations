package beans;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Created by dhruvr on 21/7/15.
 */
public class LabelAndFields {


    /**
     * list of label{@link annotations.ExcelColumn#label()} values
     */
    private final List<String> labels;

    /**
     * the name of fields
     * fields can be private , don't worry by reflection we will manage;
     */
    private final List<Field> fields;


    private LabelAndFields(Builder builder){
        this.labels = builder.labels;
        this.fields = builder.fields;
    }


    public static class Builder{
        private List<String> labels;
        private List<Field> fields;

        public Builder labels(List<String> labels){
            this.labels = labels;
            return this;
        }

        public Builder fields(List<Field> fields){
            this.fields = fields;
            return this;
        }

        public LabelAndFields build(){
            return new LabelAndFields(this);
        }
    }


    public List<String> getLabels() {
        return labels;
    }

    public List<Field> getFields() {
        return fields;
    }
}
