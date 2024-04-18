package poc;

import lombok.AllArgsConstructor;
import lombok.Data;

@AllArgsConstructor
@Data
public class Client {
    private String id;
    private String name;
    private String email;
}
